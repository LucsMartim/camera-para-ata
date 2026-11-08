import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import patch
from datetime import datetime

import sistema_ata as app


class CoreTests(unittest.TestCase):
    def test_natural_order(self):
        files = [Path('video10.mp4'), Path('video2.mp4'), Path('video1.mp4')]
        self.assertEqual([p.name for p in sorted(files, key=app.natural_key)],
                         ['video1.mp4', 'video2.mp4', 'video10.mp4'])

    def test_recording_dates(self):
        self.assertEqual(app.recording_start(Path('cam-2026-01-10-14-30-00.mp4')),
                         datetime(2026, 1, 10, 14, 30))
        self.assertIsNone(app.recording_start(Path('cam-2026-13-10-14-30-00.mp4')))
        self.assertIsNone(app.recording_start(Path('video01.mp4')))

    def test_duration_and_safe_name(self):
        self.assertEqual(app.format_duration(3661), '01:01:01')
        self.assertEqual(app.safe_name('../Reuniao / teste'), 'Reuniao_teste')
        self.assertEqual(app.safe_name('...'), 'video')

    def test_outputs_with_fictional_data(self):
        rows = [{'inicio': 0, 'fim': 4, 'arquivo': 'video_demo.mp4',
                 'texto': 'Vamos revisar os testes ficticios e registrar o resultado.',
                 'data_hora_inicio': '10/01/2026 14:30:00', 'data_hora_fim': '14:30:04'}]
        with tempfile.TemporaryDirectory() as folder, patch.object(app, 'write_docx_reports') as docx:
            root = Path(folder)
            app.write_outputs(root, [Path('video_demo.mp4')], rows, {'duracao_total': 4}, {})
            payload = json.loads((root / 'dados_processamento.json').read_text(encoding='utf-8'))
            self.assertEqual(payload['segmentos'], rows)
            self.assertIn('video_demo.mp4', (root / 'TRANSCRICAO_COMPLETA.txt').read_text(encoding='utf-8'))
            self.assertIn('RASCUNHO', (root / 'ATA_RASCUNHO.md').read_text(encoding='utf-8'))
            docx.assert_called_once()

    def test_ffmpeg_error(self):
        result = subprocess.CompletedProcess([], 1, stdout='', stderr='falha ficticia')
        with patch.object(app, 'ffmpeg_path', return_value='ffmpeg'), patch.object(app.subprocess, 'run', return_value=result):
            with self.assertRaisesRegex(RuntimeError, 'falha ficticia'):
                app.run_ffmpeg(['-version'])


if __name__ == '__main__':
    unittest.main()
