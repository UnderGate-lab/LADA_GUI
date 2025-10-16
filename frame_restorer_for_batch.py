"""
高速化版 frame_restorer.py - 最終版

スレッドセーフ修正済み：
- EOFマーカーの重複送信を防止
- クリップ順序制御の改善
- 詳細なデバッグログ
"""

import logging
import queue
import textwrap
import threading
import time
from typing import Optional
from collections import defaultdict
import heapq

import cv2
import numpy as np
import torch

from lada import LOG_LEVEL
from lada.lib import image_utils, video_utils, threading_utils, mask_utils
from lada.lib import visualization_utils
from lada.lib.mosaic_detector import MosaicDetector
from lada.lib.mosaic_detection_model import MosaicDetectionModel

logger = logging.getLogger(__name__)
logging.basicConfig(level=LOG_LEVEL)


# ====================================
# 診断クラス
# ====================================

class ParallelDiagnostics:
    """並列処理の稼働状況を診断"""
    
    def __init__(self):
        self.worker_stats = defaultdict(lambda: {
            'clips_processed': 0,
            'total_time': 0,
            'wait_time': 0,
            'gpu_time': 0,
            'last_activity': 0
        })
        self.start_time = time.time()
        self.lock = threading.Lock()
    
    def record_worker_start(self, worker_id):
        with self.lock:
            self.worker_stats[worker_id]['last_activity'] = time.time()
    
    def record_worker_processing(self, worker_id, clip_frames, processing_time):
        with self.lock:
            stats = self.worker_stats[worker_id]
            stats['clips_processed'] += 1
            stats['gpu_time'] += processing_time
            stats['last_activity'] = time.time()
    
    def record_worker_wait(self, worker_id, wait_time):
        with self.lock:
            self.worker_stats[worker_id]['wait_time'] += wait_time
    
    def get_report(self):
        with self.lock:
            total_elapsed = time.time() - self.start_time
            
            report = ["=" * 60]
            report.append("並列処理診断レポート")
            report.append("=" * 60)
            report.append(f"経過時間: {total_elapsed:.2f}秒")
            report.append("")
            
            report.append("ワーカー統計:")
            report.append("-" * 60)
            
            active_workers = 0
            total_clips = 0
            for worker_id, stats in sorted(self.worker_stats.items()):
                idle_time = time.time() - stats['last_activity']
                is_active = idle_time < 5.0
                
                if is_active:
                    active_workers += 1
                
                total_clips += stats['clips_processed']
                
                status = "🟢 稼働中" if is_active else "🔴 待機"
                report.append(
                    f"{worker_id}: {status} | "
                    f"処理: {stats['clips_processed']} | "
                    f"GPU: {stats['gpu_time']:.1f}s | "
                    f"待機: {stats['wait_time']:.1f}s"
                )
            
            report.append("")
            report.append(f"アクティブワーカー: {active_workers}/{len(self.worker_stats)}")
            report.append(f"総処理クリップ数: {total_clips}")
            
            total_gpu_time = sum(s['gpu_time'] for s in self.worker_stats.values())
            total_wait_time = sum(s['wait_time'] for s in self.worker_stats.values())
            
            if total_elapsed > 0 and len(self.worker_stats) > 0:
                gpu_utilization = (total_gpu_time / (total_elapsed * len(self.worker_stats))) * 100
                wait_ratio = (total_wait_time / (total_elapsed * len(self.worker_stats))) * 100
                
                report.append("")
                report.append(f"並列効率:")
                report.append(f"  GPU利用率: {gpu_utilization:.1f}%")
                report.append(f"  待機時間: {wait_ratio:.1f}%")
                
                if gpu_utilization < 20:
                    report.append("  ⚠️ GPU利用率が低い")
                if active_workers < len(self.worker_stats) * 0.5:
                    report.append("  ⚠️ 稼働ワーカーが少ない")
            
            report.append("=" * 60)
            
            return "\n".join(report)


# ====================================
# load_models関数
# ====================================

def load_models(device, mosaic_restoration_model_name, mosaic_restoration_model_path, 
                mosaic_restoration_config_path, mosaic_detection_model_path):
    if mosaic_restoration_model_name.startswith("deepmosaics"):
        from lada.deepmosaics.models import loadmodel, model_util
        mosaic_restoration_model = loadmodel.video(model_util.device_to_gpu_id(device), mosaic_restoration_model_path)
        pad_mode = 'reflect'
    elif mosaic_restoration_model_name.startswith("basicvsrpp"):
        from lada.basicvsrpp.inference import load_model, get_default_gan_inference_config
        if mosaic_restoration_config_path:
            config = mosaic_restoration_config_path
        else:
            config = get_default_gan_inference_config()
        mosaic_restoration_model = load_model(config, mosaic_restoration_model_path, device)
        pad_mode = 'zero'
    else:
        raise NotImplementedError()
    mosaic_detection_model = MosaicDetectionModel(mosaic_detection_model_path, device, classes=[0], conf=0.2)
    return mosaic_detection_model, mosaic_restoration_model, pad_mode


# ====================================
# 高速化版 FrameRestorer クラス
# ====================================

class FrameRestorer:
    """
    フレーム復元クラス（高速化版）
    
    既存のインターフェースを保持しつつ、内部で並列処理を実装
    """
    
    def __init__(self, device, video_file, preserve_relative_scale, max_clip_length, 
                 mosaic_restoration_model_name, mosaic_detection_model, mosaic_restoration_model, 
                 preferred_pad_mode, mosaic_detection=False,
                 parallel_clips=12,
                 enable_optimization=True,
                 enable_diagnostics=False,
                 auto_worker_count=False,
                 clip_read_timeout=15.0):
        
        self.device = device
        self.mosaic_restoration_model_name = mosaic_restoration_model_name
        self.max_clip_length = max_clip_length
        self.preserve_relative_scale = preserve_relative_scale
        self.video_meta_data = video_utils.get_video_meta_data(video_file)
        self.mosaic_detection_model = mosaic_detection_model
        self.mosaic_restoration_model = mosaic_restoration_model
        self.preferred_pad_mode = preferred_pad_mode
        self.start_ns = 0
        self.start_frame = 0
        self.mosaic_detection = mosaic_detection
        self.eof = False
        self.stop_requested = False
        
        self.enable_optimization = enable_optimization
        self.enable_diagnostics = enable_diagnostics
        self.clip_read_timeout = clip_read_timeout
        
        if auto_worker_count and torch.cuda.is_available():
            total_vram = torch.cuda.get_device_properties(0).total_memory
            self.parallel_clips = max(1, min(int((total_vram * 0.7) / (1024**3)), 8))
            logger.info(f"自動並列度計算: {self.parallel_clips} workers")
        else:
            self.parallel_clips = parallel_clips if enable_optimization else 1
        
        max_frames_in_frame_restoration_queue = (512 * 1024 * 1024) // (
            self.video_meta_data.video_width * self.video_meta_data.video_height * 3
        )
        self.frame_restoration_queue = queue.Queue(maxsize=max_frames_in_frame_restoration_queue)
        
        max_clips_in_mosaic_clips_queue = max(
            self.parallel_clips * 2,
            (512 * 1024 * 1024) // (self.max_clip_length * 256 * 256 * 4)
        )
        self.mosaic_clip_queue = queue.Queue(maxsize=max_clips_in_mosaic_clips_queue)
        
        max_clips_in_restored_clips_queue = max(
            self.parallel_clips * 2,
            (512 * 1024 * 1024) // (self.max_clip_length * 256 * 256 * 4)
        )
        self.restored_clip_queue = queue.Queue(maxsize=max_clips_in_restored_clips_queue)
        
        self.frame_detection_queue = queue.Queue()
        
        self.mosaic_detector = MosaicDetector(
            self.mosaic_detection_model, self.video_meta_data.video_file,
            frame_detection_queue=self.frame_detection_queue,
            mosaic_clip_queue=self.mosaic_clip_queue,
            device=self.device,
            max_clip_length=self.max_clip_length,
            pad_mode=self.preferred_pad_mode,
            preserve_relative_scale=self.preserve_relative_scale,
            dont_preserve_relative_scale=(not self.preserve_relative_scale)
        )
        
        if self.enable_optimization and self.parallel_clips > 1:
            self.clip_restoration_threads = []
            self.clip_ordering_thread = None
            self.unordered_clips_queue = queue.Queue(maxsize=max_clips_in_restored_clips_queue * 2)
            self.next_expected_clip_id = 0
            self.clip_counter = 0
            self.clip_counter_lock = threading.Lock()
            self.eof_marker_sent = False
            self.eof_marker_lock = threading.Lock()
            self.workers_finished_count = 0  # 終了したワーカー数
            self.workers_finished_lock = threading.Lock()
            logger.info(f"並列処理モード有効: {self.parallel_clips} workers")
        else:
            self.clip_restoration_thread = None
            logger.info("シングルスレッドモード")
        
        self.frame_restoration_thread = None
        self.clip_restoration_threads_should_be_running = False
        self.frame_restoration_thread_should_be_running = False
        self.clip_ordering_thread_should_be_running = False
        
        self.queue_stats = {
            "restored_clip_queue_max_size": 0,
            "restored_clip_queue_wait_time_put": 0,
            "restored_clip_queue_wait_time_get": 0,
            "mosaic_clip_queue_wait_time_get": 0,
            "frame_restoration_queue_max_size": 0,
            "frame_restoration_queue_wait_time_get": 0,
            "frame_restoration_queue_wait_time_put": 0,
            "frame_detection_queue_wait_time_get": 0,
            "parallel_clips_processed": 0,
            "clip_timeout_count": 0,
        }
        
        self.consecutive_timeouts = 0
        self.max_consecutive_timeouts = 3
        self.last_timeout_frame = -1
        
        if self.enable_diagnostics:
            self.diagnostics = ParallelDiagnostics()
            logger.info("✅ 診断機能を有効化")
        else:
            self.diagnostics = None
    
    def _get_next_clip_id(self):
        with self.clip_counter_lock:
            clip_id = self.clip_counter
            self.clip_counter += 1
            return clip_id
    
    def start(self, start_ns=0):
        if self.enable_optimization and self.parallel_clips > 1:
            self._start_parallel_mode(start_ns)
        else:
            self._start_single_mode(start_ns)
    
    def _start_single_mode(self, start_ns):
        assert self.frame_restoration_thread is None and (not hasattr(self, 'clip_restoration_thread') or self.clip_restoration_thread is None)
        
        self.start_ns = start_ns
        self.start_frame = video_utils.offset_ns_to_frame_num(self.start_ns, self.video_meta_data.video_fps_exact)
        self.stop_requested = False
        self.frame_restoration_thread_should_be_running = True
        self.clip_restoration_threads_should_be_running = True
        
        self.frame_restoration_thread = threading.Thread(target=self._frame_restoration_worker)
        self.clip_restoration_thread = threading.Thread(target=self._clip_restoration_worker_single)
        
        self.mosaic_detector.start(start_ns=start_ns)
        self.clip_restoration_thread.start()
        self.frame_restoration_thread.start()
    
    def _start_parallel_mode(self, start_ns):
        assert self.frame_restoration_thread is None and len(self.clip_restoration_threads) == 0
        
        self.start_ns = start_ns
        self.start_frame = video_utils.offset_ns_to_frame_num(self.start_ns, self.video_meta_data.video_fps_exact)
        self.stop_requested = False
        self.next_expected_clip_id = 0
        self.clip_counter = 0
        self.eof_marker_sent = False
        self.workers_finished_count = 0  # リセット
        
        self.frame_restoration_thread_should_be_running = True
        self.clip_restoration_threads_should_be_running = True
        self.clip_ordering_thread_should_be_running = True
        
        self.mosaic_detector.start(start_ns=start_ns)
        
        for i in range(self.parallel_clips):
            thread = threading.Thread(
                target=self._clip_restoration_worker_parallel,
                name=f"ClipWorker-{i}"
            )
            thread.start()
            self.clip_restoration_threads.append(thread)
        
        self.clip_ordering_thread = threading.Thread(
            target=self._clip_ordering_worker,
            name="ClipOrderingWorker"
        )
        self.clip_ordering_thread.start()
        
        self.frame_restoration_thread = threading.Thread(
            target=self._frame_restoration_worker,
            name="FrameRestorationWorker"
        )
        self.frame_restoration_thread.start()
        
        logger.info(f"起動完了: {self.parallel_clips} workers")
    
    def stop(self):
        if self.enable_optimization and self.parallel_clips > 1:
            self._stop_parallel_mode()
        else:
            self._stop_single_mode()
    
    def _stop_single_mode(self):
        logger.debug("FrameRestorer: stopping...")
        start = time.time()
        self.stop_requested = True
        self.clip_restoration_threads_should_be_running = False
        self.frame_restoration_thread_should_be_running = False
        
        self.mosaic_detector.stop()
        
        threading_utils.put_closing_queue_marker(self.mosaic_clip_queue, "mosaic_clip_queue")
        threading_utils.empty_out_queue(self.restored_clip_queue, "restored_clip_queue")
        if hasattr(self, 'clip_restoration_thread') and self.clip_restoration_thread:
            self.clip_restoration_thread.join()
            logger.debug("clip restoration worker: stopped")
        self.clip_restoration_thread = None
        
        threading_utils.put_closing_queue_marker(self.frame_detection_queue, "frame_detection_queue")
        threading_utils.put_closing_queue_marker(self.restored_clip_queue, "restored_clip_queue")
        threading_utils.empty_out_queue(self.frame_restoration_queue, "frame_restoration_queue")
        if self.frame_restoration_thread:
            self.frame_restoration_thread.join()
            logger.debug("frame restoration worker: stopped")
        self.frame_restoration_thread = None
        
        threading_utils.empty_out_queue(self.mosaic_clip_queue, "mosaic_clip_queue")
        threading_utils.empty_out_queue(self.restored_clip_queue, "restored_clip_queue")
        threading_utils.empty_out_queue(self.frame_detection_queue, "frame_detection_queue")
        threading_utils.empty_out_queue(self.frame_restoration_queue, "frame_restoration_queue")
        
        logger.debug(f"FrameRestorer: stopped, took {time.time() - start}")
    
    def _stop_parallel_mode(self):
        logger.debug("FrameRestorer (parallel): stopping...")
        start_time = time.time()
        
        if self.diagnostics is not None:
            import sys
            print("\n" + "="*70, file=sys.stdout)
            print("📊 並列処理診断レポート", file=sys.stdout)
            print("="*70, file=sys.stdout)
            try:
                report = self.diagnostics.get_report()
                print(report, file=sys.stdout)
                sys.stdout.flush()
            except Exception as e:
                print(f"⚠️ レポート生成エラー: {e}", file=sys.stderr)
            print("="*70 + "\n", file=sys.stdout)
            sys.stdout.flush()
        
        self.stop_requested = True
        self.clip_restoration_threads_should_be_running = False
        self.frame_restoration_thread_should_be_running = False
        self.clip_ordering_thread_should_be_running = False
        
        self.mosaic_detector.stop()
        
        for _ in range(self.parallel_clips):
            threading_utils.put_closing_queue_marker(self.mosaic_clip_queue, "mosaic_clip_queue")
        
        threading_utils.empty_out_queue(self.unordered_clips_queue, "unordered_clips_queue")
        
        for thread in self.clip_restoration_threads:
            if thread:
                thread.join(timeout=2.0)
        self.clip_restoration_threads.clear()
        logger.debug("clip restoration workers: stopped")
        
        threading_utils.put_closing_queue_marker(self.unordered_clips_queue, "unordered_clips_queue")
        threading_utils.empty_out_queue(self.restored_clip_queue, "restored_clip_queue")
        
        if self.clip_ordering_thread:
            self.clip_ordering_thread.join(timeout=2.0)
            self.clip_ordering_thread = None
        logger.debug("clip ordering worker: stopped")
        
        threading_utils.put_closing_queue_marker(self.frame_detection_queue, "frame_detection_queue")
        threading_utils.put_closing_queue_marker(self.restored_clip_queue, "restored_clip_queue")
        threading_utils.empty_out_queue(self.frame_restoration_queue, "frame_restoration_queue")
        
        if self.frame_restoration_thread:
            self.frame_restoration_thread.join(timeout=2.0)
            self.frame_restoration_thread = None
        logger.debug("frame restoration worker: stopped")
        
        if torch.cuda.is_available():
            torch.cuda.empty_cache()
        
        threading_utils.empty_out_queue(self.mosaic_clip_queue, "mosaic_clip_queue")
        threading_utils.empty_out_queue(self.restored_clip_queue, "restored_clip_queue")
        threading_utils.empty_out_queue(self.frame_detection_queue, "frame_detection_queue")
        threading_utils.empty_out_queue(self.frame_restoration_queue, "frame_restoration_queue")
        threading_utils.empty_out_queue(self.unordered_clips_queue, "unordered_clips_queue")
        
        logger.debug(f"FrameRestorer: stopped, took {time.time() - start_time:.2f}s")
        logger.info(f"📈 並列クリップ処理数: {self.queue_stats.get('parallel_clips_processed', 0)}")
    
    def _clip_restoration_worker_single(self):
        logger.debug("clip restoration worker: started")
        eof = False
        while self.clip_restoration_threads_should_be_running:
            s = time.time()
            clip = self.mosaic_clip_queue.get()
            self.queue_stats["mosaic_clip_queue_wait_time_get"] += time.time() - s
            if self.stop_requested:
                logger.debug("clip restoration worker: mosaic_clip_queue consumer unblocked")
            if clip is None:
                if not self.stop_requested:
                    eof = True
                    self.clip_restoration_threads_should_be_running = False
                    self.queue_stats["restored_clip_queue_max_size"] = max(self.restored_clip_queue.qsize()+1, self.queue_stats["restored_clip_queue_max_size"])
                    s = time.time()
                    self.restored_clip_queue.put(None)
                    self.queue_stats["restored_clip_queue_wait_time_put"] += time.time() -s
                    logger.debug("clip restoration worker: restored_clip_queue producer unblocked")
            else:
                self._restore_clip(clip)
                self.queue_stats["restored_clip_queue_max_size"] = max(self.restored_clip_queue.qsize()+1, self.queue_stats["restored_clip_queue_max_size"])
                s = time.time()
                self.restored_clip_queue.put(clip)
                self.queue_stats["restored_clip_queue_wait_time_put"] += time.time() - s
                if self.stop_requested:
                    logger.debug("clip restoration worker: restored_clip_queue producer unblocked")
        if eof:
            logger.debug("clip restoration worker: stopped itself, EOF")
    
    def _clip_restoration_worker_parallel(self):
        worker_name = threading.current_thread().name
        logger.debug(f"{worker_name}: 起動")
        
        error_count = 0
        max_errors = 10
        clips_processed = 0
        
        while self.clip_restoration_threads_should_be_running:
            try:
                wait_start = time.time()
                clip = self.mosaic_clip_queue.get(timeout=0.5)
                wait_time = time.time() - wait_start
                
                if self.diagnostics:
                    self.diagnostics.record_worker_wait(worker_name, wait_time)
                
                if clip is None:
                    logger.info(f"{worker_name}: EOFマーカー受信 (処理数: {clips_processed})")
                    
                    # EOFマーカーを他のワーカーのために戻す
                    self.mosaic_clip_queue.put(None)
                    
                    # 終了カウントをインクリメント
                    with self.workers_finished_lock:
                        self.workers_finished_count += 1
                        finished_count = self.workers_finished_count
                        logger.info(f"{worker_name}: 終了 ({finished_count}/{self.parallel_clips})")
                        
                        # 最後のワーカーだけがEOFマーカーを送信
                        if finished_count == self.parallel_clips:
                            logger.info(f"{worker_name}: 全ワーカー終了 - EOFマーカー送信")
                            self.unordered_clips_queue.put(None)
                    
                    break
                
                clip_id = self._get_next_clip_id()
                clip_length = len(clip.get_clip_images())
                
                logger.debug(f"{worker_name}: クリップ {clip_id} 処理開始 (frames={clip_length})")
                
                if self.diagnostics:
                    self.diagnostics.record_worker_start(worker_name)
                
                gpu_start = time.time()
                
                try:
                    self._restore_clip(clip)
                    
                    gpu_time = time.time() - gpu_start
                    
                    if self.diagnostics:
                        self.diagnostics.record_worker_processing(worker_name, clip_length, gpu_time)
                    
                    self.unordered_clips_queue.put((clip_id, clip))
                    self.queue_stats['parallel_clips_processed'] += 1
                    clips_processed += 1
                    error_count = 0
                    
                    logger.debug(f"{worker_name}: クリップ {clip_id} 完了 (GPU時間={gpu_time:.2f}s)")
                    
                except Exception as e:
                    logger.error(f"{worker_name}: クリップ {clip_id} 処理エラー: {e}", exc_info=True)
                    error_count += 1
                    if error_count >= max_errors:
                        logger.error(f"{worker_name}: エラー多発のため終了")
                        break
            
            except queue.Empty:
                if self.stop_requested:
                    break
                continue
        
        logger.debug(f"{worker_name}: 終了 (合計: {clips_processed})")
    
    def _clip_ordering_worker(self):
        pending_clips = []
        eof = False
        last_output_time = time.time()
        stall_warning_interval = 10.0
        
        while self.clip_ordering_thread_should_be_running:
            try:
                item = self.unordered_clips_queue.get(timeout=0.1)
                
                if item is None:
                    if not self.stop_requested:
                        eof = True
                        logger.debug(
                            f"clip_ordering_worker: EOF受信\n"
                            f"  pending_clips: {len(pending_clips)}\n"
                            f"  next_expected_id: {self.next_expected_clip_id}\n"
                            f"  total_processed: {self.queue_stats['parallel_clips_processed']}"
                        )
                        break
                    else:
                        break
                
                clip_id, clip = item
                heapq.heappush(pending_clips, (clip_id, clip))
                logger.info(
                    f"clip_ordering_worker: クリップ {clip_id} 受信\n"
                    f"  pending_clips IDs: {sorted([c[0] for c in pending_clips])}\n"
                    f"  next_expected: {self.next_expected_clip_id}"
                )
                
                while pending_clips and pending_clips[0][0] == self.next_expected_clip_id:
                    _, ordered_clip = heapq.heappop(pending_clips)
                    logger.info(f"clip_ordering_worker: クリップ {self.next_expected_clip_id} 送出")
                    self.restored_clip_queue.put(ordered_clip)
                    self.next_expected_clip_id += 1
                    last_output_time = time.time()
                
                if pending_clips and (time.time() - last_output_time) > stall_warning_interval:
                    logger.warning(
                        f"⚠️ クリップ順序制御が停滞:\n"
                        f"  次期待ID: {self.next_expected_clip_id}\n"
                        f"  保留中クリップ: {len(pending_clips)}\n"
                        f"  最小ID: {pending_clips[0][0]}\n"
                        f"  最大ID: {max(c[0] for c in pending_clips)}"
                    )
                    last_output_time = time.time()
            
            except queue.Empty:
                if self.stop_requested:
                    break
                continue
        
        if eof:
            logger.info(
                f"clip_ordering_worker: EOF処理開始\n"
                f"  pending_clips: {len(pending_clips)}\n"
                f"  pending_clips IDs: {sorted([c[0] for c in pending_clips])}\n"
                f"  next_expected_id: {self.next_expected_clip_id}\n"
                f"  total_sent: {self.next_expected_clip_id}\n"
                f"  total_generated: {self.queue_stats['parallel_clips_processed']}"
            )
            
            while pending_clips:
                expected_id = self.next_expected_clip_id
                actual_id = pending_clips[0][0]
                
                logger.info(
                    f"clip_ordering_worker: EOF処理ループ\n"
                    f"  期待ID: {expected_id}\n"
                    f"  実際ID: {actual_id}\n"
                    f"  残りpending: {sorted([c[0] for c in pending_clips])}"
                )
                
                if actual_id == expected_id:
                    _, ordered_clip = heapq.heappop(pending_clips)
                    self.restored_clip_queue.put(ordered_clip)
                    logger.info(f"clip_ordering_worker: クリップ {expected_id} 送出（EOF処理中）")
                    self.next_expected_clip_id += 1
                
                elif actual_id > expected_id:
                    error_msg = (
                        f"❌ クリップ欠落検出\n"
                        f"   期待ID: {expected_id}\n"
                        f"   実際ID: {actual_id}\n"
                        f"   欠落数: {actual_id - expected_id}\n"
                        f"   pending IDs: {sorted([c[0] for c in pending_clips])}\n"
                        f"   送出済み: {expected_id}個\n"
                        f"   生成済み: {self.queue_stats['parallel_clips_processed']}個"
                    )
                    logger.error(error_msg)
                    raise RuntimeError(f"クリップID {expected_id} が欠落しました")
                
                else:
                    logger.error(f"❌ 異常: 既に送出済みのID {actual_id} が残っています")
                    heapq.heappop(pending_clips)
            
            logger.info(f"clip_ordering_worker: 全クリップ送出完了 (total={self.next_expected_clip_id})")
            self.restored_clip_queue.put(None)
            logger.info("clip_ordering_worker: EOFマーカー送出完了")
    
    def _restore_clip_frames(self, images):
        if self.mosaic_restoration_model_name.startswith("deepmosaics"):
            from lada.deepmosaics.inference import restore_video_frames
            from lada.deepmosaics.models import model_util
            restored_clip_images = restore_video_frames(
                model_util.device_to_gpu_id(self.device),
                self.mosaic_restoration_model,
                images
            )
        elif self.mosaic_restoration_model_name.startswith("basicvsrpp"):
            from lada.basicvsrpp.inference import inference
            restored_clip_images = inference(
                self.mosaic_restoration_model,
                images,
                self.device
            )
        else:
            raise NotImplementedError()
        return restored_clip_images
    
    def _restore_frame(self, frame, frame_num, restored_clips):
        for buffered_clip in [c for c in restored_clips if c.frame_start == frame_num]:
            clip_img, clip_mask, orig_clip_box, orig_crop_shape, pad_after_resize = buffered_clip.pop()
            clip_img = image_utils.unpad_image(clip_img, pad_after_resize)
            clip_mask = image_utils.unpad_image(clip_mask, pad_after_resize)
            clip_img = image_utils.resize(clip_img, orig_crop_shape[:2])
            clip_mask = image_utils.resize(clip_mask, orig_crop_shape[:2], interpolation=cv2.INTER_NEAREST)
            t, l, b, r = orig_clip_box
            blend_mask = mask_utils.create_blend_mask(clip_mask)
            blended_img = (
                frame[t:b + 1, l:r + 1, :] * (1 - blend_mask[..., None]) +
                clip_img * (blend_mask[..., None])
            ).clip(0, 255).astype(np.uint8)
            frame[t:b + 1, l:r + 1, :] = blended_img
    
    def _restore_clip(self, clip):
        if self.mosaic_detection:
            restored_clip_images = visualization_utils.draw_mosaic_detections(clip)
        else:
            images = clip.get_clip_images()
            restored_clip_images = self._restore_clip_frames(images)
        
        assert len(restored_clip_images) == len(clip.get_clip_images())
        
        for i in range(len(restored_clip_images)):
            assert clip.data[i][0].shape == restored_clip_images[i].shape
            clip.data[i] = (
                restored_clip_images[i],
                clip.data[i][1],
                clip.data[i][2],
                clip.data[i][3],
                clip.data[i][4]
            )
    
    def _collect_garbage(self, clip_buffer):
        processed_clips = list(filter(lambda _clip: len(_clip) == 0, clip_buffer))
        for processed_clip in processed_clips:
            clip_buffer.remove(processed_clip)
    
    def _contains_at_least_one_clip_starting_after_frame_num(self, frame_num, clip_buffer):
        return len(clip_buffer) > 0 and frame_num < max(clip_buffer, key=lambda c: c.frame_start).frame_start
    
    def _read_next_frame(self, video_frames_generator, expected_frame_num):
        try:
            frame, frame_pts = next(video_frames_generator)
        except StopIteration:
            s = time.time()
            elem = self.frame_detection_queue.get()
            self.queue_stats["frame_detection_queue_wait_time_get"] += time.time() - s
            if self.stop_requested:
                logger.debug("frame restoration worker: frame_detection_queue consumer unblocked")
            assert elem is None
            return None
        
        s = time.time()
        elem = self.frame_detection_queue.get()
        self.queue_stats["frame_detection_queue_wait_time_get"] += time.time() - s
        if self.stop_requested:
            logger.debug("frame restoration worker: frame_detection_queue consumer unblocked")
            return None
        assert elem is not None
        detection_frame_num, mosaic_detected = elem
        assert self.stop_requested or detection_frame_num == expected_frame_num
        return mosaic_detected, frame, frame_pts
    
    def _read_next_clip(self, current_frame_num, clip_buffer):
        if self.enable_optimization and self.parallel_clips > 1:
            try:
                clip = self.restored_clip_queue.get(timeout=self.clip_read_timeout)
                
                self.consecutive_timeouts = 0
                self.last_timeout_frame = -1
                
                if self.stop_requested or clip is None:
                    return False
                clip_buffer.append(clip)
                return True
            except queue.Empty:
                self.queue_stats["clip_timeout_count"] += 1
                
                if self.last_timeout_frame == current_frame_num:
                    self.consecutive_timeouts += 1
                else:
                    self.consecutive_timeouts = 1
                    self.last_timeout_frame = current_frame_num
                
                logger.error(
                    f"⚠️ クリップ読み取りタイムアウト (frame={current_frame_num}, 連続{self.consecutive_timeouts}回)\n"
                    f"  📊 キュー状態診断:\n"
                    f"    mosaic_clip_queue: {self.mosaic_clip_queue.qsize()}/{self.mosaic_clip_queue.maxsize}\n"
                    f"    restored_clip_queue: {self.restored_clip_queue.qsize()}/{self.restored_clip_queue.maxsize}\n"
                    f"    unordered_clips_queue: {self.unordered_clips_queue.qsize()}/{self.unordered_clips_queue.maxsize}\n"
                    f"    frame_restoration_queue: {self.frame_restoration_queue.qsize()}/{self.frame_restoration_queue.maxsize}\n"
                    f"  🔢 処理状況:\n"
                    f"    並列クリップ処理数: {self.queue_stats['parallel_clips_processed']}\n"
                    f"    次期待クリップID: {self.next_expected_clip_id}\n"
                    f"    クリップカウンター: {self.clip_counter}\n"
                    f"    現在のclip_buffer数: {len(clip_buffer)}\n"
                    f"  ⚠️ 処理が停止している可能性があります"
                )
                
                if self.consecutive_timeouts >= self.max_consecutive_timeouts:
                    error_msg = (
                        f"❌ 致命的エラー: クリップ読み取り失敗 (frame={current_frame_num})\n"
                        f"   {self.max_consecutive_timeouts}回連続タイムアウト - デッドロックの可能性\n"
                        f"   バッチ処理のため、正確性を保つために処理を中断します"
                    )
                    logger.error(error_msg)
                    
                    if self.diagnostics:
                        print("\n" + "="*70)
                        print("🔴 緊急診断レポート（デッドロック検出）")
                        print("="*70)
                        print(self.diagnostics.get_report())
                        print("="*70 + "\n")
                    
                    raise RuntimeError(error_msg)
                
                return True
        else:
            s = time.time()
            clip = self.restored_clip_queue.get()
            self.queue_stats["restored_clip_queue_wait_time_get"] += time.time() - s
            if self.stop_requested:
                logger.debug("frame restoration worker: restored_clip_queue consumer unblocked")
            if clip is None:
                return False
            assert self.stop_requested or clip.frame_start >= current_frame_num
            clip_buffer.append(clip)
            return True
    
    def _frame_restoration_worker(self):
        logger.debug("frame restoration worker: started")
        with video_utils.VideoReader(self.video_meta_data.video_file) as video_reader:
            if self.start_ns > 0:
                video_reader.seek(self.start_ns)
            
            video_frames_generator = video_reader.frames()
            frame_num = self.start_frame
            clips_remaining = True
            clip_buffer = []
            
            while self.frame_restoration_thread_should_be_running:
                _frame_result = self._read_next_frame(video_frames_generator, frame_num)
                if _frame_result is None:
                    if not self.stop_requested:
                        self.eof = True
                        self.frame_restoration_thread_should_be_running = False
                        self.frame_restoration_queue.put(None)
                    break
                else:
                    mosaic_detected, frame, frame_pts = _frame_result
                
                if mosaic_detected:
                    while clips_remaining and not self._contains_at_least_one_clip_starting_after_frame_num(
                        frame_num, clip_buffer
                    ):
                        clips_remaining = self._read_next_clip(frame_num, clip_buffer)
                    
                    self._restore_frame(frame, frame_num, clip_buffer)
                    self.queue_stats["frame_restoration_queue_max_size"] = max(
                        self.frame_restoration_queue.qsize()+1, 
                        self.queue_stats["frame_restoration_queue_max_size"]
                    )
                    s = time.time()
                    self.frame_restoration_queue.put((frame, frame_pts))
                    self.queue_stats["frame_restoration_queue_wait_time_put"] += time.time() - s
                    if self.stop_requested:
                        logger.debug("frame restoration worker: frame_restoration_queue producer unblocked")
                    self._collect_garbage(clip_buffer)
                else:
                    self.queue_stats["frame_restoration_queue_max_size"] = max(
                        self.frame_restoration_queue.qsize()+1, 
                        self.queue_stats["frame_restoration_queue_max_size"]
                    )
                    s = time.time()
                    self.frame_restoration_queue.put((frame, frame_pts))
                    self.queue_stats["frame_restoration_queue_wait_time_put"] += time.time() - s
                    if self.stop_requested:
                        logger.debug("frame restoration worker: frame_restoration_queue producer unblocked")
                
                frame_num += 1
            
            if self.eof:
                logger.debug("frame restoration worker: stopped itself, EOF")
    
    def __iter__(self):
        return self
    
    def __next__(self):
        if self.eof and self.frame_restoration_queue.empty():
            raise StopIteration
        else:
            while not self.stop_requested:
                s = time.time()
                elem = self.frame_restoration_queue.get()
                self.queue_stats["frame_restoration_queue_wait_time_get"] += time.time() - s
                if self.stop_requested:
                    logger.debug("frame_restoration_queue consumer unblocked")
                if elem is None and not self.stop_requested:
                    raise StopIteration
                return elem
    
    def get_frame_restoration_queue(self):
        return self.frame_restoration_queue
