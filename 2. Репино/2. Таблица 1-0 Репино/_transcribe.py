import sys
from pathlib import Path
from faster_whisper import WhisperModel

AUDIO = Path(__file__).parent / "Таблица 1-0 Репино.m4a"
OUT_TXT = AUDIO.with_suffix(".txt")
OUT_SRT = AUDIO.with_suffix(".srt")

print(f"Loading model 'medium' (downloads on first use) ...", flush=True)
model = WhisperModel("medium", device="cpu", compute_type="int8")
print("Model loaded. Transcribing...", flush=True)

segments, info = model.transcribe(
    str(AUDIO),
    language="ru",
    beam_size=5,
    vad_filter=True,
    vad_parameters=dict(min_silence_duration_ms=500),
)
print(f"Detected language={info.language} (p={info.language_probability:.2f}), duration={info.duration:.1f}s", flush=True)

txt_lines = []
srt_lines = []
def ts(s):
    h = int(s // 3600); s -= h*3600
    m = int(s // 60); s -= m*60
    sec = int(s); ms = int((s-sec)*1000)
    return f"{h:02d}:{m:02d}:{sec:02d},{ms:03d}"

for i, seg in enumerate(segments, 1):
    line = seg.text.strip()
    txt_lines.append(line)
    srt_lines.append(f"{i}\n{ts(seg.start)} --> {ts(seg.end)}\n{line}\n")
    print(f"[{seg.start:6.1f}s] {line}", flush=True)

OUT_TXT.write_text("\n".join(txt_lines), encoding="utf-8")
OUT_SRT.write_text("\n".join(srt_lines), encoding="utf-8")
print(f"\nWrote: {OUT_TXT.name}\nWrote: {OUT_SRT.name}", flush=True)
