"""Transcripción de llamadas con marcas de tiempo y separación de hablantes.

Pasos: decodificar el audio -> detectar quién habla cuándo (diarización) ->
transcribir cada tramo con Whisper -> armar el texto "[mm:ss] Persona N: ...".

Uso por línea de comandos:
    python transcribir.py llamada.wav [--hablantes 2] [--modelo medium]
Escribe llamada.txt al lado del audio.
"""
import argparse
from pathlib import Path

import av
import numpy as np
import sherpa_onnx

MODELOS = Path(__file__).parent / "modelos"
SR = 16000
MAX_TRAMO = 25.0   # Whisper procesa ventanas de hasta 30 s
HUECO_MAX = 1.0    # une tramos del mismo hablante separados por menos de esto
MIN_TRAMO = 1.0    # tramos más cortos que esto se alargan para que Whisper los entienda


def cargar_audio(ruta):
    """Devuelve el audio en mono, 16 kHz, float32 (cualquier formato que lea ffmpeg)."""
    with av.open(str(ruta)) as cont:
        rs = av.AudioResampler(format="flt", layout="mono", rate=SR)
        trozos = []
        for frame in cont.decode(audio=0):
            for f in rs.resample(frame):
                trozos.append(f.to_ndarray().reshape(-1))
        for f in rs.resample(None):
            trozos.append(f.to_ndarray().reshape(-1))
    return np.concatenate(trozos).astype(np.float32)


def _diarizador(hablantes):
    M = MODELOS
    cfg = sherpa_onnx.OfflineSpeakerDiarizationConfig(
        segmentation=sherpa_onnx.OfflineSpeakerSegmentationModelConfig(
            pyannote=sherpa_onnx.OfflineSpeakerSegmentationPyannoteModelConfig(
                model=str(M / "sherpa-onnx-pyannote-segmentation-3-0" / "model.int8.onnx")
            ),
            num_threads=4,
        ),
        embedding=sherpa_onnx.SpeakerEmbeddingExtractorConfig(
            model=str(M / "embedding.onnx"), num_threads=4
        ),
        clustering=sherpa_onnx.FastClusteringConfig(num_clusters=hablantes),
        min_duration_on=0.3,
        min_duration_off=0.5,
    )
    return sherpa_onnx.OfflineSpeakerDiarization(cfg)


def _reconocedor(modelo):
    d = MODELOS / f"sherpa-onnx-whisper-{modelo}"
    return sherpa_onnx.OfflineRecognizer.from_whisper(
        encoder=str(d / f"{modelo}-encoder.int8.onnx"),
        decoder=str(d / f"{modelo}-decoder.int8.onnx"),
        tokens=str(d / f"{modelo}-tokens.txt"),
        language="es",
        task="transcribe",
        num_threads=4,
    )


def _agrupar(segmentos):
    """Une tramos consecutivos del mismo hablante y parte los que pasan de MAX_TRAMO."""
    grupos = []
    for s in segmentos:
        if grupos and grupos[-1]["hablante"] == s.speaker and s.start - grupos[-1]["fin"] < HUECO_MAX \
                and s.end - grupos[-1]["inicio"] <= MAX_TRAMO:
            grupos[-1]["fin"] = s.end
        else:
            grupos.append({"inicio": s.start, "fin": s.end, "hablante": s.speaker})
    salida = []
    for g in grupos:
        t = g["inicio"]
        while g["fin"] - t > MAX_TRAMO:
            salida.append({**g, "inicio": t, "fin": t + MAX_TRAMO})
            t += MAX_TRAMO
        salida.append({**g, "inicio": t})
    return salida


def transcribir(ruta, hablantes=2, modelo="medium", progreso=None):
    """Devuelve una lista de intervenciones: {inicio, fin, hablante (1..N), texto}."""
    informar = progreso or (lambda *_: None)
    if not (MODELOS / f"sherpa-onnx-whisper-{modelo}").exists():
        raise FileNotFoundError(
            f"Falta el modelo '{modelo}'. Ejecutá: python descargar_modelos.py {modelo}"
        )
    informar("Leyendo el audio...", 0.02)
    audio = cargar_audio(ruta)
    if len(audio) < SR:
        raise ValueError("El audio es demasiado corto.")

    informar("Identificando quién habla en cada momento...", 0.05)
    segs = _diarizador(hablantes).process(audio).sort_by_start_time()
    tramos = _agrupar(segs)

    rec = _reconocedor(modelo)
    # Se numera a las personas por orden de aparición.
    orden = {}
    resultado = []
    for i, t in enumerate(tramos):
        informar("Transcribiendo...", 0.15 + 0.85 * i / max(len(tramos), 1))
        ini, fin = t["inicio"], t["fin"]
        if fin - ini < MIN_TRAMO:
            extra = (MIN_TRAMO - (fin - ini)) / 2
            ini, fin = max(0, ini - extra), min(len(audio) / SR, fin + extra)
        a = audio[int(max(0, ini - 0.1) * SR): int(fin * SR) + int(0.1 * SR)]
        st = rec.create_stream()
        st.accept_waveform(SR, a)
        rec.decode_stream(st)
        texto = st.result.text.strip()
        if not texto:
            continue
        n = orden.setdefault(t["hablante"], len(orden) + 1)
        resultado.append({"inicio": t["inicio"], "fin": t["fin"], "hablante": n, "texto": texto})
    informar("Listo", 1.0)
    return resultado


def marca(segundos):
    s = int(segundos)
    return f"{s // 60:02d}:{s % 60:02d}"


def a_texto(intervenciones):
    return "\n".join(
        f"[{marca(i['inicio'])}] Persona {i['hablante']}: {i['texto']}" for i in intervenciones
    ) + "\n"


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("audio")
    p.add_argument("--hablantes", type=int, default=2)
    p.add_argument("--modelo", default="medium", choices=["small", "medium"])
    a = p.parse_args()
    r = transcribir(a.audio, a.hablantes, a.modelo, lambda m, f: print(f"{f:4.0%} {m}", end="\r"))
    destino = Path(a.audio).with_suffix(".txt")
    destino.write_text(a_texto(r), encoding="utf-8")
    print(f"\nGuardado en {destino}")


if __name__ == "__main__":
    main()
