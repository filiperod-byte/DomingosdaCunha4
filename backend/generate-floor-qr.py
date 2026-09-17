"""Regerar os QR locais. Dependência: qrcode==8.2 (pip install qrcode==8.2)."""
import json
from pathlib import Path
import qrcode
import qrcode.image.svg
root = Path(__file__).resolve().parents[1]
config = json.loads((root / 'config.json').read_text())
output = root / 'assets' / 'qr-floors'
output.mkdir(parents=True, exist_ok=True)
for floor in config['building']['floors']:
    n = floor['floor']
    url = f'https://filiperod-byte.github.io/DomingosdaCunha4/general-report.html?floor={n}'
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    qr.make_image(image_factory=qrcode.image.svg.SvgPathImage).save(output / f'piso-{n}.svg')
