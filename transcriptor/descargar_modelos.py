"""Descarga (una sola vez) los modelos que usa el transcriptor.

Todo corre en la propia computadora: el audio no se manda a ningún servicio
externo. Los modelos se guardan en la carpeta ./modelos.

Uso:
    python descargar_modelos.py            # modelo "medium" (recomendado)
    python descargar_modelos.py small      # más liviano y rápido, menos preciso
"""
import sys
import tarfile
import urllib.request
from pathlib import Path

BASE = "https://github.com/k2-fsa/sherpa-onnx/releases/download"
DESTINO = Path(__file__).parent / "modelos"

COMUN = [
    ("asr-models/silero_vad.onnx", "silero_vad.onnx"),
    (
        "speaker-recongition-models/"
        "3dspeaker_speech_eres2net_base_sv_zh-cn_3dspeaker_16k.onnx",
        "embedding.onnx",
    ),
]
SEGMENTACION = "speaker-segmentation-models/sherpa-onnx-pyannote-segmentation-3-0.tar.bz2"


def bajar(url, destino):
    print(f"Descargando {url.split('/')[-1]} ...")
    urllib.request.urlretrieve(url, destino)


def extraer(tar, carpeta):
    with tarfile.open(tar) as t:
        t.extractall(carpeta, filter="data")
    tar.unlink()


def main():
    tam = sys.argv[1] if len(sys.argv) > 1 else "medium"
    if tam not in ("small", "medium"):
        sys.exit("El modelo debe ser 'small' o 'medium'.")
    DESTINO.mkdir(exist_ok=True)

    for ruta, nombre in COMUN:
        if not (DESTINO / nombre).exists():
            bajar(f"{BASE}/{ruta}", DESTINO / nombre)

    if not (DESTINO / "sherpa-onnx-pyannote-segmentation-3-0").exists():
        tar = DESTINO / "seg.tar.bz2"
        bajar(f"{BASE}/{SEGMENTACION}", tar)
        extraer(tar, DESTINO)

    if not (DESTINO / f"sherpa-onnx-whisper-{tam}").exists():
        tar = DESTINO / "whisper.tar.bz2"
        bajar(f"{BASE}/asr-models/sherpa-onnx-whisper-{tam}.tar.bz2", tar)
        extraer(tar, DESTINO)
    print("Listo.")


if __name__ == "__main__":
    main()
