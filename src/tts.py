#!/usr/bin/env python3
"""src/narration/*.txt を Piper (ローカル・無料のTTS) で読み上げ、src/audio/slide-N.wav を作る

generate.py が書き出した narration/slide-N.txt を総なめにして、対応する
audio/slide-N.wav を生成するだけの単純なスクリプト。既存の .wav は毎回上書きする。

同じく generate.py が書き出す narration/slide-N.json（{"speaker": ..., "speed": ...}）が
あれば、話者・速度としてそれを使う。speaker は音声モデルの speaker_id_map（例: female/male）
のキー、speed は 1.0 が標準で大きいほど速い（内部で Piper の length_scale = 1/speed に変換）。

生成した音声は、次に generate.py を実行すればそのスライドに自動で埋め込まれる
（README.md の「実装メモ」を参照）。
"""

import json
import sys
import wave
from pathlib import Path

from piper import PiperVoice
from piper.voice import SynthesisConfig

VOICE_PATH = Path(__file__).parent / "voices" / "ja_JA-hi_fi_captain-medium.onnx"


def load_syn_config(txt_path: Path, voice: PiperVoice) -> SynthesisConfig:
    """対応する slide-N.json（speaker/speed）を SynthesisConfig に変換する。なければデフォルト。"""
    json_path = txt_path.with_suffix(".json")
    if not json_path.exists():
        return SynthesisConfig()

    data = json.loads(json_path.read_text())
    speaker_id_map = voice.config.speaker_id_map or {}
    speaker_id = speaker_id_map.get(data.get("speaker"))

    speed = data.get("speed") or 1.0
    length_scale = 1.0 / speed if speed else None

    return SynthesisConfig(speaker_id=speaker_id, length_scale=length_scale)


def ensure_voice_model() -> None:
    """事前条件チェック: 音声モデルが無ければ入手方法を案内して終了する。"""
    if not VOICE_PATH.exists():
        print(f"音声モデルが見つかりません: {VOICE_PATH}")
        print("先に以下でダウンロードしてください:")
        print(f"  .venv/bin/python -m piper.download_voices --download-dir {VOICE_PATH.parent} ja_JA-hi_fi_captain-medium")
        sys.exit(1)


def collect_narration_files(narration_dir: Path) -> list[Path]:
    """事前条件チェック: ナレーション原稿(.txt)を集める。1件も無ければ案内して終了する。"""
    txt_files = sorted(narration_dir.glob("slide-*.txt"))
    if not txt_files:
        print(f"ナレーション原稿が見つかりません: {narration_dir}/slide-*.txt")
        print("先に generate.py を実行して narration/*.txt を作ってください。")
        sys.exit(1)
    return txt_files


def synthesize_all(txt_files: list[Path], voice: PiperVoice, audio_dir: Path) -> None:
    """本処理: 原稿ごとに音声合成し、audio_dir に slide-N.wav として書き出す。"""
    audio_dir.mkdir(exist_ok=True)
    for txt_path in txt_files:
        text = txt_path.read_text().strip()
        if not text:
            continue
        syn_config = load_syn_config(txt_path, voice)
        wav_path = audio_dir / f"{txt_path.stem}.wav"
        with wave.open(str(wav_path), "wb") as wav_file:
            voice.synthesize_wav(text, wav_file, syn_config=syn_config)
        print(f"wrote {wav_path} (speaker_id={syn_config.speaker_id}, length_scale={syn_config.length_scale})")


def main() -> None:
    # パス準備
    root = Path(__file__).parent
    narration_dir = root / "narration"
    audio_dir = root / "audio"

    # 事前条件チェック（無ければここで終了する）
    ensure_voice_model()
    txt_files = collect_narration_files(narration_dir)

    # 音声モデルを読み込み、全原稿を合成する
    voice = PiperVoice.load(str(VOICE_PATH))
    synthesize_all(txt_files, voice, audio_dir)


if __name__ == "__main__":
    main()
