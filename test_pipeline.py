import cv2
import torch
import numpy as np
import time
import threading
import queue
import win32com.client
import pythoncom
import json
import os
from ultralytics import YOLO

# Vosk와 PyAudio는 선택 사항으로 처리 (설치 환경 문제 대비)
VOSK_AVAILABLE = False
PYAUDIO_AVAILABLE = False
try:
    from vosk import Model, KaldiRecognizer
    import pyaudio
    VOSK_AVAILABLE = True
    PYAUDIO_AVAILABLE = True
except ImportError:
    pass

# ==========================================
# Optimized MVP Test Pipeline: TTS (음성 안내) 버전
# ==========================================

class MVPTestPipeline:
    def __init__(self):
        print("음성 지원 모드로 전환 중... 모델 로딩 중...")
        
        # 설정값
        self.inference_size = (320, 320)
        self.frame_skip = 3
        self.frame_count = 0
        self.K_DEPTH = 3000.0 
        self.running = False  # 제어용 플래그

        # 음성 안내 설정 (볼륨 및 뮤트)
        self.volume = 100  # 0 ~ 100
        self.is_muted = False

        # TTS 큐 및 스레드 초기화
        self.speech_queue = queue.Queue()
        self.tts_thread = threading.Thread(target=self._tts_worker, daemon=True)
        self.tts_thread.start()

        # STT (음성 인식) 초기화
        self.model_path = "model-ko" # 한국어 모델 폴더명
        self.stt_thread = None
        if VOSK_AVAILABLE and PYAUDIO_AVAILABLE and os.path.exists(self.model_path):
            try:
                self.stt_thread = threading.Thread(target=self._stt_worker, daemon=True)
                self.stt_thread.start()
            except Exception as e:
                print(f"⚠️ STT 초기화 중 오류 발생: {e}")
        else:
            reason = "모델 폴더 없음" if not os.path.exists(self.model_path) else "라이브러리(vosk/pyaudio) 미설치"
            print(f"⚠️ 음성 명령이 비활성화되었습니다. (원인: {reason})")
        
        # 시작 알림 (스피커 확인용)
        self.speak("시스템을 시작합니다.")

        # 음성 상태 관리
        self.announced_objects = {} # {label: last_seen_time}
        self.announce_timeout = 8.0 # 8초 동안 안 보이면 안내 목록에서 삭제 (다시 나타나면 말함)

        # 모델 로딩
        self.yolo_model = YOLO('yolov8n.pt') 
        self.depth_model_type = "MiDaS_small"
        self.midas = torch.hub.load("intel-isl/MiDaS", self.depth_model_type, trust_repo=True)
        self.device = torch.device("cuda") if torch.cuda.is_available() else torch.device("cpu")
        self.midas.to(self.device).eval()
        
        midas_transforms = torch.hub.load("intel-isl/MiDaS", "transforms", trust_repo=True)
        self.transform = midas_transforms.small_transform if self.depth_model_type == "MiDaS_small" else midas_transforms.dpt_transform

        self.last_objects = []
        self.last_depth_map = None
        self.last_depth_viz = None
        
        # 웹 스트리밍용 버퍼
        self.last_web_frame = None
        self.frame_lock = threading.Lock()

        # 한국어 클래스 맵
        self.class_names_ko = {
            'person': '사람', 'bicycle': '자전거', 'car': '자동차', 'motorcycle': '오토바이',
            'bus': '버스', 'truck': '트럭', 'traffic light': '신호등', 'stop sign': '정지 표지판',
            'bench': '벤치', 'dog': '개', 'cat': '고양이', 'backpack': '배낭', 'umbrella': '우산',
            'handbag': '핸드백', 'tie': '넥타이', 'suitcase': '여행가방', 'sports ball': '공',
            'bottle': '병', 'wine glass': '와인잔', 'cup': '컵', 'fork': '포크', 'knife': '칼',
            'spoon': '숟가락', 'bowl': '그릇', 'banana': '바나나', 'apple': '사과', 'sandwich': '샌드위치',
            'orange': '오렌지', 'broccoli': '브로콜리', 'carrot': '당근', 'hot dog': '핫도그', 'pizza': '피자',
            'donut': '도넛', 'cake': '케이크', 'chair': '의자', 'couch': '소파', 'potted plant': '화분',
            'bed': '침대', 'dining table': '식탁', 'toilet': '변기', 'tv': 'TV', 'laptop': '노트북',
            'mouse': '마우스', 'remote': '리모컨', 'keyboard': '키보드', 'cell phone': '핸드폰',
            'microwave': '전자레인지', 'oven': '오븐', '토스터': '토스터', 'sink': '싱크대',
            'refrigerator': '냉장고', 'book': '책', 'clock': '시계', 'vase': '꽃병', 'scissors': '가위',
            'teddy bear': '곰인형', 'hair drier': '헤어드라이어', 'toothbrush': '칫솔'
        }

        # Walking assistance ROI (Center 40%)
        self.roi_x_min = 0.3
        self.roi_x_max = 0.7

    def _tts_worker(self):
        """별도 스레드에서 SAPI 엔진을 초기화하고 안내를 처리 (가장 확실한 윈도우 방식)"""
        pythoncom.CoInitialize()
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        
        while True:
            # 큐에서 데이터를 가져옴 (텍스트, 강제중지여부)
            item = self.speech_queue.get()
            if item is None: break
            
            text, force_stop = item
            
            # 뮤트 상태면 무시 (단, 강제 종료 안내는 예외)
            if self.is_muted and not force_stop:
                self.speech_queue.task_done()
                continue

            # 실시간 볼륨 적용
            speaker.Volume = self.volume

            # force_stop이 True이면 현재 말하고 있는 것과 밀려있는 큐를 모두 무시하고 즉시 말함
            # SAPI Flag: 2 (SVSFPurgeBeforeSpeak)
            flags = 2 if force_stop else 0
            
            print(f"[TTS 발화 시작] {text} (강제종료: {force_stop})")
            try:
                speaker.Speak(text, flags)
            except Exception as e:
                print(f"[TTS 오류] {e}")
            print(f"[TTS 발화 완료] {text}")
            self.speech_queue.task_done()

    def _stt_worker(self):
        """마이크 소리를 듣고 명령어를 인식하는 스레드"""
        model = Model(self.model_path)
        rec = KaldiRecognizer(model, 16000)
        p = pyaudio.PyAudio()
        stream = p.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8000)
        stream.start_stream()

        print("🎙️ 음성 인식 준비 완료. 명령을 기다립니다...")

        while True:
            data = stream.read(4000, exception_on_overflow=False)
            if rec.AcceptWaveform(data):
                result = json.loads(rec.Result())
                text = result.get("text", "").replace(" ", "")
                if not text: continue

                print(f"👂 음성 인식 결과: {text}")
                self.handle_command(text)

    def handle_command(self, text):
        """음성 인식을 통해 들어온 텍스트를 분석하여 명령 수행"""
        # 명령어 판별 (공백 제거 후 비교)
        text = text.replace(" ", "")
        
        if "종료" in text:
            self.speak("시스템을 종료합니다.", force_stop=True)
            self.running = False
        elif "다시시작" in text or "다시실행" in text:
            self.speak("시스템을 다시 시작합니다.", force_stop=True)
        elif "볼륨올려" in text:
            self.volume = min(100, self.volume + 20)
            self.speak(f"볼륨을 올렸습니다. 현재 볼륨 {self.volume}")
        elif "볼륨내려" in text:
            self.volume = max(0, self.volume - 20)
            self.speak(f"볼륨을 내렸습니다. 현재 볼륨 {self.volume}")
        elif "조용히해" in text or "정지해" in text:
            self.is_muted = True
            self.speak("음성 안내를 일시 정지합니다.", force_stop=True)
        elif "말해줘" in text or "다시말해" in text:
            self.is_muted = False
            self.speak("음성 안내를 다시 시작합니다.")

    def speak(self, text, force_stop=False):
        """안내 문구를 큐에 추가 (비동기)"""
        if force_stop:
            # 기존 큐에 쌓인 모든 메시지 무시하도록 큐 비우기 시도
            while not self.speech_queue.empty():
                try:
                    self.speech_queue.get_nowait()
                    self.speech_queue.task_done()
                except:
                    break
        self.speech_queue.put((text, force_stop))

    def stage2_yolo_optimized(self, frame):
        results = self.yolo_model(frame, imgsz=320, verbose=False) 
        objects = []
        for r in results:
            boxes = r.boxes
            for box in boxes:
                b = box.xyxy[0].cpu().numpy().astype(int)
                cls_id = int(box.cls[0])
                model_label = self.yolo_model.names[cls_id]
                ko_label = self.class_names_ko.get(model_label, model_label)
                objects.append({'box': b, 'label': ko_label})
        return objects

    def stage3_depth_optimized(self, frame):
        small_frame = cv2.resize(frame, (256, 256)) 
        img = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)
        input_batch = self.transform(img).to(self.device)

        with torch.no_grad():
            prediction = self.midas(input_batch)
            prediction = torch.nn.functional.interpolate(
                prediction.unsqueeze(1),
                size=frame.shape[:2],
                mode="bicubic",
                align_corners=False,
            ).squeeze()

        depth_map = prediction.cpu().numpy()
        depth_min, depth_max = depth_map.min(), depth_map.max()
        depth_norm = (255 * (depth_map - depth_min) / (depth_max - depth_min + 1e-5)).astype(np.uint8)
        depth_color = cv2.applyColorMap(depth_norm, cv2.COLORMAP_MAGMA)
        
        return depth_map, depth_color

    def raw_to_meters(self, raw_val):
        if raw_val <= 0: return float('inf')
        meters = self.K_DEPTH / (raw_val + 1e-5)
        return meters

    def run(self):
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("카메라를 열 수 없습니다.")
            return

        window_name_main = "MVP Test - Color (YOLO)"
        window_name_depth = "MVP Test - Depth (MiDaS)"
        cv2.namedWindow(window_name_main)
        cv2.namedWindow(window_name_depth)

        print("\n=== 음성 안내(TTS)가 최적화된 MVP 파이프라인 시작 ===")
        
        # 시작 시 안내 음성 추가 (웹에서 다시 시작할 때도 나옴)
        self.speak("보조 시스템 안내를 시작합니다.", force_stop=True)

        self.running = True
        last_log_time = 0
        log_interval = 6.0 

        while self.running:
            ret, frame = cap.read()
            if not ret: break
            
            self.frame_count += 1
            current_time = time.time()
            
            # --- 파이프라인 연산 ---
            if self.frame_count % self.frame_skip == 1 or self.last_depth_map is None:
                self.last_objects = self.stage2_yolo_optimized(frame)
                self.last_depth_map, self.last_depth_viz = self.stage3_depth_optimized(frame)
            
            display_frame = frame.copy()
            should_log = (current_time - last_log_time) >= log_interval

            # --- ROI 필터링 및 가장 가까운 물체 선택 ---
            h, w = frame.shape[:2]
            roi_left = int(w * self.roi_x_min)
            roi_right = int(w * self.roi_x_max)
            
            closest_obj = None
            min_meters = float('inf')

            for obj in self.last_objects:
                b = obj['box']
                cx = (b[0] + b[2]) // 2
                cy = int(b[3] * 0.9)
                
                # ROI 내부에 중심이 있는 경우에만 처리
                if roi_left <= cx <= roi_right:
                    h_d, w_d = self.last_depth_map.shape
                    cx_d, cy_d = max(0, min(cx, w_d-1)), max(0, min(cy, h_d-1))
                    
                    raw_val = self.last_depth_map[cy_d, cx_d]
                    meters = self.raw_to_meters(raw_val)
                    
                    # 가장 가까운 물체 갱신
                    if meters < min_meters:
                        min_meters = meters
                        closest_obj = {
                            'label': obj['label'],
                            'box': b,
                            'meters': meters,
                            'cx': cx
                        }

            # --- 시각화 및 안내 ---
            # ROI 가이드 라인 표시
            cv2.line(display_frame, (roi_left, 0), (roi_left, h), (0, 0, 255), 2)
            cv2.line(display_frame, (roi_right, 0), (roi_right, h), (0, 0, 255), 2)

            current_labels = set()
            if closest_obj and min_meters < 10.0:
                b = closest_obj['box']
                label_name = closest_obj['label']
                meters = closest_obj['meters']
                current_labels.add(label_name)

                # 시각화 (선택된 물체만 강조)
                cv2.rectangle(display_frame, (b[0], b[1]), (b[2], b[3]), (0, 0, 255), 3)
                cv2.putText(display_frame, f"TARGET: {label_name} {meters:.1f}m", (b[0], b[1]-10), 
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)

                # --- 음성 안내 로직 ---
                if label_name not in self.announced_objects:
                    self.speak(f"전방에 {label_name}가 있습니다. 거리는 {meters:.1f} 미터입니다.")
                    self.announced_objects[label_name] = current_time

                if should_log:
                    print(f"[보행 보조] 장애물 감지: {label_name} | 거리: {meters:.1f}m")

            # 안내 상태 업데이트 (오랫동안 안 보인 사물은 목록에서 제거)
            for label in list(self.announced_objects.keys()):
                if label not in current_labels:
                    if current_time - self.announced_objects[label] > self.announce_timeout:
                        del self.announced_objects[label]

            if should_log:
                last_log_time = current_time

            # 화면 표시

            # 화면 표시
            # 웹 스트리밍용으로 현재 프레임 저장
            with self.frame_lock:
                self.last_web_frame = display_frame.copy()

            cv2.imshow(window_name_main, display_frame)
            if self.last_depth_viz is not None:
                cv2.imshow(window_name_depth, self.last_depth_viz)
            
            # 종료 로직
            key = cv2.waitKey(1) & 0xFF
            if key == ord('q') or key == 27: # Q나 ESC
                break
            
            # 창이 닫혔는지 확인
            if cv2.getWindowProperty(window_name_main, cv2.WND_PROP_VISIBLE) < 1:
                break

        cap.release()
        cv2.destroyAllWindows()

if __name__ == "__main__":
    pipeline = MVPTestPipeline()
    pipeline.run()
