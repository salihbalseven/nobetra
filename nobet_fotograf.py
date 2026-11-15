"""Read duty-grid cells locally with Windows' Turkish OCR engine."""

import base64
from bisect import bisect_right
from datetime import date, timedelta
from io import BytesIO
import json
import os
import re
from pathlib import Path
import subprocess

from PIL import Image, ImageOps


OCR = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapPixelFormat,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapAlphaMode,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and
        $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } |
    Select-Object -First 1
function Await($operation, [Type]$type) {
    $task = $asTask.MakeGenericMethod($type).Invoke($null, @($operation))
    return $task.GetAwaiter().GetResult()
}
$bytes = [Convert]::FromBase64String([Console]::In.ReadToEnd())
$memory = [IO.MemoryStream]::new($bytes)
$stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($memory)
try {
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $image = Await ($decoder.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
        [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore)) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('tr'))
        if ($null -eq $engine) { throw 'Windows Türkçe OCR desteği bulunamadı.' }
        $result = Await ($engine.RecognizeAsync($image)) ([Windows.Media.Ocr.OcrResult])
        $words = @(
            foreach ($line in $result.Lines) {
                foreach ($word in $line.Words) {
                    $box = $word.BoundingRect
                    @{text=$word.Text; x=($box.X+$box.Width/2); y=($box.Y+$box.Height/2)}
                }
            }
        )
        ConvertTo-Json -InputObject $words -Compress
    } finally { $image.Dispose() }
} finally { $stream.Dispose(); $memory.Dispose() }
"""


def open_image(path):
    with Image.open(path) as source:
        if source.format not in ('PNG', 'JPEG'):
            raise ValueError('PNG, JPG veya JPEG fotoğraf seçin.')
        image = ImageOps.exif_transpose(source).convert('RGB')
    image.thumbnail((4800, 4800), Image.Resampling.LANCZOS)
    return image


def recognize(image):
    if os.name != 'nt':
        raise ValueError('Fotoğraf okuma için Windows Türkçe OCR desteği gerekli.')
    powershell = Path(os.environ['SystemRoot']) / 'System32/WindowsPowerShell/v1.0/powershell.exe'
    with BytesIO() as buffer:
        image.save(buffer, format='PNG')
        payload = base64.b64encode(buffer.getvalue()).decode('ascii')
    script = base64.b64encode(OCR.encode('utf-16-le')).decode('ascii')
    try:
        result = subprocess.run([str(powershell), '-NoProfile', '-NonInteractive', '-EncodedCommand', script],
                                input=payload, capture_output=True, encoding='utf-8', errors='replace',
                                timeout=50, creationflags=subprocess.CREATE_NO_WINDOW)
    except subprocess.TimeoutExpired as error:
        raise ValueError('Fotoğraf okuma zaman aşımına uğradı.') from error
    if result.returncode:
        detail = result.stderr.strip()
        if 'Türkçe OCR' in detail:
            raise ValueError('Windows Türkçe OCR desteği bulunamadı.')
        raise ValueError('Fotoğraf okunamadı. Görseli yeniden seçip köşeleri kontrol edin.')
    return json.loads(result.stdout)


def perspective(width, height, corners):
    """Find the inverse projective transform from the output rectangle to the image."""
    equations = []
    for (x, y), (u, v) in zip(((0, 0), (width, 0), (width, height), (0, height)), corners):
        equations.extend(([x, y, 1, 0, 0, 0, -u*x, -u*y, u],
                          [0, 0, 0, x, y, 1, -v*x, -v*y, v]))
    for col in range(8):
        pivot = max(range(col, 8), key=lambda row: abs(equations[row][col]))
        equations[col], equations[pivot] = equations[pivot], equations[col]
        value = equations[col][col]
        if abs(value) < 1e-9:
            raise ValueError('Dört farklı çizelge köşesini sırayla işaretleyin.')
        equations[col] = [number / value for number in equations[col]]
        for row in range(8):
            if row != col:
                factor = equations[row][col]
                equations[row] = [a - factor*b for a, b in zip(equations[row], equations[col])]
    return [row[-1] for row in equations]


def read_cells(image, corners, template, last_row, addresses, date_groups=()):
    widths, heights = template['widths'], template['heights'][:last_row]
    scale = min(2, 3600 / max(sum(widths), sum(heights)))
    width, height = round(sum(widths) * scale), round(sum(heights) * scale)
    straight = image.transform((width, height), Image.Transform.PERSPECTIVE,
                               perspective(width, height, corners), Image.Resampling.BICUBIC)
    x_edges = [sum(widths[:i]) * scale for i in range(1, len(widths) + 1)]
    y_edges = [sum(heights[:i]) * scale for i in range(1, len(heights) + 1)]
    definitions = {(cell['col'] - 1, cell['row'] - 1): cell['address']
                   for cell in template['cells'] if cell['address'] in addresses}
    collected = {}
    for word in recognize(ImageOps.autocontrast(straight.convert('L'), cutoff=1)):
        position = bisect_right(x_edges, word['x']), bisect_right(y_edges, word['y'])
        address = definitions.get(position)
        if address:
            collected.setdefault(address, []).append(word)
    values = {
        key: ' '.join(word['text'] for word in sorted(words, key=lambda w: (round(w['y'] / 12), w['x'])))
        for key, words in collected.items()
    }
    if not date_groups:
        return values
    cells = {cell['address']: cell for cell in template['cells']}
    def crop_date(address):
        cell = cells[address]
        left, top = round(sum(widths[:cell['col']-1]) * scale), round(sum(heights[:cell['row']-1]) * scale)
        right, bottom = round(x_edges[cell['endcol']-1]), round(y_edges[cell['endrow']-1])
        crop = straight.crop((left + 2, top + 2, right - 2, bottom - 2))
        if cell['rotation'] in (90, 180):
            crop = crop.transpose(Image.Transpose.ROTATE_270)
        return crop
    dates = []
    for addresses in date_groups:
        candidates = set()
        for address, day_offset in addresses:
            crop = ImageOps.autocontrast(crop_date(address).convert('L'), cutoff=1)
            if crop.height < 120:
                factor = min(3, 120 / crop.height)
                crop = crop.resize((round(crop.width * factor), round(crop.height * factor)), Image.Resampling.LANCZOS)
            if cells[address]['rotation']:
                options = (crop, crop.transpose(Image.Transpose.ROTATE_180))
            else:
                options = (crop, crop.point(lambda pixel: 0 if pixel < 120 else 255))
            for candidate in options:
                words = recognize(ImageOps.expand(candidate, border=12, fill=255))
                found = list(printed_dates(' '.join(word['text'] for word in words)))
                candidates.update((date.fromisoformat(day) - timedelta(days=day_offset)).isoformat() for day in found)
                if found:
                    break
            if len(candidates) == 1:
                break
        dates.append(next(iter(candidates)) if len(candidates) == 1 else '')
    return values, dates


def printed_dates(text):
    for day, month, year in re.findall(r'(?<!\d)(\d{1,2})\s*[./,-]\s*(\d{1,2})\s*[./,-]?\s*(20\d{2})(?!\d)', text):
        try:
            yield date(int(year), int(month), int(day)).isoformat()
        except ValueError:
            continue
