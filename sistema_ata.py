from __future__ import annotations

import json
import os
import pathlib
import re
import subprocess
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk


ROOT = pathlib.Path(__file__).resolve().parent
MODELS = ROOT / ".modelos"


def natural_key(path: pathlib.Path) -> list[object]:
    return [int(value) if value.isdigit() else value.lower() for value in re.split(r"(\d+)", path.name)]


def ffmpeg_path() -> str:
    import imageio_ffmpeg

    return imageio_ffmpeg.get_ffmpeg_exe()


def duration_seconds(path: pathlib.Path) -> float:
    import av

    with av.open(str(path)) as container:
        return float(container.duration / av.time_base) if container.duration else 0.0


def format_duration(seconds: float) -> str:
    total = int(round(seconds))
    hours, rest = divmod(total, 3600)
    minutes, secs = divmod(rest, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def safe_name(value: str) -> str:
    cleaned = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("_.")
    return cleaned or "video"


def recording_start(path: pathlib.Path) -> datetime | None:
    match = re.search(
        r"(?<!\d)(20\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})[-_](\d{2})(?!\d)",
        path.name,
    )
    if not match:
        return None
    try:
        return datetime(*map(int, match.groups()))
    except ValueError:
        return None


def format_recording_time(value: datetime | None) -> str:
    return value.strftime("%d/%m/%Y %H:%M:%S") if value else "não identificado"


def run_ffmpeg(arguments: list[str]) -> None:
    completed = subprocess.run(
        [ffmpeg_path(), "-hide_banner", "-loglevel", "error", *arguments],
        capture_output=True,
        text=True,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    if completed.returncode:
        raise RuntimeError(completed.stderr.strip() or "Falha ao converter o áudio.")


def transcribe(audio: pathlib.Path, duration: float, progress_callback, model_name: str = "medium") -> tuple[list[dict], str, float]:
    from faster_whisper import WhisperModel

    model = WhisperModel(model_name, device="cpu", compute_type="int8", download_root=str(MODELS))
    segments, info = model.transcribe(
        str(audio), language="pt", beam_size=5, vad_filter=True, word_timestamps=False
    )
    rows = []
    for item in segments:
        if item.text.strip():
            rows.append({"inicio": round(item.start, 2), "fim": round(item.end, 2), "texto": item.text.strip()})
        progress_callback(item.end, duration)
    return rows, info.language, info.language_probability


def add_table_borders(table) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:color"), "D9D9D9")
        borders.append(element)
    table._tbl.tblPr.append(borders)


def shade_cell(cell, fill: str) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margin(cell, value: int = 110) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    margins = cell._tc.get_or_add_tcPr().first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        cell._tc.get_or_add_tcPr().append(margins)
    for edge in ("top", "left", "bottom", "right"):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)


def prepare_document(title: str, meeting: dict):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt, RGBColor

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.7)
    section.left_margin = Cm(2.1)
    section.right_margin = Cm(2.1)

    styles = doc.styles
    styles["Normal"].font.name = "Aptos"
    styles["Normal"].font.size = Pt(10.5)
    for style_name in ("Title", "Heading 1", "Heading 2"):
        styles[style_name].font.name = "Aptos Display"
        styles[style_name].font.color.rgb = RGBColor(0, 0, 0)
    title_style_ppr = styles["Title"]._element.get_or_add_pPr()
    title_border = title_style_ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr")
    if title_border is not None:
        title_style_ppr.remove(title_border)

    logo = meeting.get("logo", "")
    if logo and pathlib.Path(logo).is_file():
        paragraph = doc.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        try:
            paragraph.add_run().add_picture(logo, height=Cm(1.8))
        except Exception:
            paragraph.add_run("Logomarca não pôde ser inserida")
    heading = doc.add_paragraph(style="Title")
    heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
    heading.add_run(title)
    intro = doc.add_paragraph(
        "Este documento registra as informações fornecidas sobre a reunião e organiza o conteúdo transcrito para conferência, decisões e acompanhamento."
    )
    intro.paragraph_format.space_after = Pt(12)

    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    fields = (
        ("Cliente", meeting.get("cliente", "Não informado")),
        ("Reunião", meeting.get("titulo", "Não informado")),
        ("Data e horário", f"{meeting.get('data', 'Não informado')} {meeting.get('horario', '')}".strip()),
        ("Local", meeting.get("local", "Não informado")),
        ("Participantes", meeting.get("participantes", "Não informado")),
    )
    for index, (label, value) in enumerate(fields):
        cells = table.add_row().cells
        cells[0].width = Cm(4.0)
        cells[1].width = Cm(12.5)
        cells[0].text = label
        cells[1].text = value or "Não informado"
        shade_cell(cells[0], "245371")
        cells[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        cells[0].paragraphs[0].runs[0].bold = True
        if index % 2:
            shade_cell(cells[1], "EEF3F4")
        set_cell_margin(cells[0])
        set_cell_margin(cells[1])
    add_table_borders(table)
    return doc


def write_docx_reports(output: pathlib.Path, videos: list[pathlib.Path], rows: list[dict], metadata: dict, meeting: dict) -> None:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    transcript_doc = prepare_document("Relatório detalhado da transcrição", meeting)
    transcript_doc.add_heading("Resumo do processamento", level=1)
    transcript_doc.add_paragraph(
        f"Foram processados {len(videos)} arquivo(s), totalizando {format_duration(metadata['duracao_total'])}. "
        "Cada fala abaixo mantém horário e vídeo de origem para facilitar a conferência no áudio."
    )
    transcript_doc.add_heading("Arquivos de origem", level=1)
    for index, video in enumerate(videos, 1):
        transcript_doc.add_paragraph(f"{index}. {video.name}")
    transcript_doc.add_heading("Transcrição cronológica", level=1)
    for row in rows:
        paragraph = transcript_doc.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(5)
        marker = paragraph.add_run(
            f"{row.get('data_hora_inicio', 'Horário não identificado')} a {row.get('data_hora_fim', '')}\n"
            f"Posição {format_duration(row['inicio'])} a {format_duration(row['fim'])} | {row['arquivo']}\n"
        )
        marker.bold = True
        paragraph.add_run(row["texto"])
    transcript_doc.add_heading("Observações para revisão", level=1)
    transcript_doc.add_paragraph(
        meeting.get("observacoes") or "Conferir nomes próprios, termos técnicos, decisões, responsáveis e prazos diretamente no áudio."
    )
    transcript_doc.sections[0].footer.paragraphs[0].text = "Relatório gerado pelo Camera para ATA"
    transcript_doc.sections[0].footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    transcript_doc.save(output / "RELATORIO_TRANSCRICAO.docx")

    ata_doc = prepare_document("Ata da reunião", meeting)
    complete_text = " ".join(row["texto"] for row in rows)
    paragraphs = [
        item.strip() for item in re.split(r"(?<=[.!?])\s+", complete_text)
        if len(item.strip()) >= 45 and not item.strip().endswith("...")
    ]
    actions = [
        item for item in paragraphs
        if any(term in item.lower() for term in ("precisa", "deve", "vamos", "ficou", "responsável", "prazo", "registrar"))
    ]
    ata_doc.add_heading("Objetivo e contexto", level=1)
    ata_doc.add_paragraph(meeting.get("observacoes") or "Reunião registrada por gravação e organizada para revisão dos participantes.")
    ata_doc.add_heading("Assuntos discutidos", level=1)
    for item in (paragraphs[:12] or ["Nenhum assunto pôde ser identificado com segurança."]):
        ata_doc.add_paragraph(item, style="List Bullet")
    ata_doc.add_heading("Decisões e encaminhamentos", level=1)
    for item in (actions[:10] or ["Nenhum encaminhamento foi identificado automaticamente com segurança."]):
        ata_doc.add_paragraph(item, style="List Bullet")
    ata_doc.add_heading("Responsáveis e prazos", level=1)
    ata_doc.add_paragraph("Revisar e completar este campo antes da aprovação da ata.")
    ata_doc.add_heading("Validação", level=1)
    ata_doc.add_paragraph(
        "Esta ata foi preparada a partir de transcrição automática. Os participantes devem revisar nomes, termos técnicos, decisões, responsáveis e prazos antes da aprovação."
    )
    ata_doc.add_paragraph("\nResponsável pela revisão: __________________________________________")
    ata_doc.add_paragraph("Data da aprovação: ____/____/________")
    ata_doc.add_heading("Registro cronológico para conferência", level=1)
    ata_doc.add_paragraph(
        "Os horários abaixo foram calculados a partir da data e da hora inicial presentes no nome de cada arquivo."
    )
    for row in rows:
        paragraph = ata_doc.add_paragraph()
        paragraph.paragraph_format.space_after = Pt(4)
        run = paragraph.add_run(
            f"{row.get('data_hora_inicio', 'Horário não identificado')} a {row.get('data_hora_fim', '')} | {row['arquivo']}\n"
        )
        run.bold = True
        paragraph.add_run(row["texto"])
    ata_doc.sections[0].footer.paragraphs[0].text = "Ata gerada pelo Camera para ATA"
    ata_doc.sections[0].footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    ata_doc.save(output / "ATA_REUNIAO.docx")


def write_outputs(output: pathlib.Path, videos: list[pathlib.Path], rows: list[dict], metadata: dict, meeting: dict) -> None:
    transcript = "\n".join(
        f"[{row.get('data_hora_inicio', 'horário não identificado')} - {row.get('data_hora_fim', '')}] "
        f"[posição {format_duration(row['inicio'])} - {format_duration(row['fim'])}] "
        f"[{row['arquivo']}] {row['texto']}"
        for row in rows
    )
    (output / "TRANSCRICAO_COMPLETA.txt").write_text(transcript, encoding="utf-8")
    (output / "dados_processamento.json").write_text(
        json.dumps({**metadata, "videos": [str(item) for item in videos], "segmentos": rows}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    topics = []
    actions = []
    for row in rows:
        text = row["texto"].strip()
        lowered = text.lower()
        if len(text) >= 35 and text not in topics:
            topics.append(text)
        if any(term in lowered for term in ("precisa", "deve", "vamos", "ficou", "responsável", "prazo", "registrar")):
            actions.append(text)

    date_label = datetime.now().strftime("%d/%m/%Y %H:%M")
    ata = [
        "# ATA DA REUNIÃO — RASCUNHO PARA REVISÃO",
        "",
        f"**Processada em:** {date_label}  ",
        f"**Quantidade de vídeos:** {len(videos)}  ",
        f"**Duração total:** {format_duration(metadata['duracao_total'])}  ",
        "**Situação:** conteúdo gerado automaticamente; exige conferência humana.",
        "",
        "## Arquivos processados",
        "",
        *[f"- {item.name}" for item in videos],
        "",
        "## Principais trechos identificados",
        "",
        *([f"- {item}" for item in topics[:12]] or ["- Nenhum trecho suficientemente claro foi identificado."]),
        "",
        "## Possíveis encaminhamentos — conferir no áudio",
        "",
        *([f"- {item}" for item in actions[:10]] or ["- Nenhum encaminhamento foi identificado com segurança."]),
        "",
        "## Decisões, responsáveis e prazos",
        "",
        "Preencher após a revisão da transcrição e dos áudios. O sistema não completa informações ausentes por suposição.",
        "",
        "## Arquivos para conferência",
        "",
        "- `AUDIO_REUNIAO_COMPLETA.mp3`: áudio contínuo para o cliente.",
        "- Pasta `audios_por_video`: áudio separado de cada gravação.",
        "- `TRANSCRICAO_COMPLETA.txt`: falas com horário e vídeo de origem.",
    ]
    (output / "ATA_RASCUNHO.md").write_text("\n".join(ata), encoding="utf-8")
    write_docx_reports(output, videos, rows, metadata, meeting)


class AtaApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Camera para ATA")
        self.geometry("960x680")
        self.geometry("1060x860")
        self.minsize(900, 760)
        self.configure(bg="#eef3f4")
        self.files: list[pathlib.Path] = []
        self.output_var = tk.StringVar(value=str(ROOT / "reunioes_processadas"))
        self.status_var = tk.StringVar(value="Selecione os vídeos da reunião.")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_text_var = tk.StringVar(value="0%")
        self.client_var = tk.StringVar()
        self.title_var = tk.StringVar()
        self.date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        self.time_var = tk.StringVar()
        self.location_var = tk.StringVar()
        self.logo_var = tk.StringVar()
        self.started_at: float | None = None
        self.current_status = "Selecione os vídeos da reunião."
        self.current_progress = 0.0
        self._build()

    def _build(self) -> None:
        header = tk.Frame(self, bg="#245371", padx=28, pady=20)
        header.pack(fill="x")
        tk.Label(header, text="Camera para ATA", bg="#245371", fg="white", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tk.Label(
            header,
            text="Organize os vídeos de 30 minutos, extraia os áudios e gere uma única transcrição.",
            bg="#245371",
            fg="#eef3f4",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(4, 0))

        body = tk.Frame(self, bg="#eef3f4", padx=28, pady=22)
        body.pack(fill="both", expand=True)

        info = tk.LabelFrame(body, text=" 1  Informações da reunião ", bg="#eef3f4", fg="#19384d", font=("Segoe UI", 11, "bold"), padx=12, pady=10)
        info.pack(fill="x", pady=(0, 14))
        fields = (
            ("Cliente", self.client_var, 0, 0),
            ("Título da reunião", self.title_var, 0, 2),
            ("Data", self.date_var, 1, 0),
            ("Horário", self.time_var, 1, 2),
            ("Local", self.location_var, 2, 0),
        )
        for label, variable, row_index, column in fields:
            tk.Label(info, text=label, bg="#eef3f4", fg="#29495e", font=("Segoe UI", 9, "bold")).grid(row=row_index, column=column, sticky="w", padx=(0, 8), pady=4)
            entry = tk.Entry(info, textvariable=variable, font=("Segoe UI", 10))
            entry.grid(row=row_index, column=column + 1, sticky="ew", padx=(0, 16), pady=4, ipady=4)
        info.columnconfigure(1, weight=1)
        info.columnconfigure(3, weight=1)
        tk.Label(info, text="Participantes", bg="#eef3f4", fg="#29495e", font=("Segoe UI", 9, "bold")).grid(row=3, column=0, sticky="nw", pady=4)
        self.participants_text = tk.Text(info, height=2, font=("Segoe UI", 10), wrap="word")
        self.participants_text.grid(row=3, column=1, columnspan=3, sticky="ew", padx=(0, 16), pady=4)
        tk.Label(info, text="Observações", bg="#eef3f4", fg="#29495e", font=("Segoe UI", 9, "bold")).grid(row=4, column=0, sticky="nw", pady=4)
        self.notes_text = tk.Text(info, height=2, font=("Segoe UI", 10), wrap="word")
        self.notes_text.grid(row=4, column=1, columnspan=3, sticky="ew", padx=(0, 16), pady=4)
        tk.Label(info, text="Logomarca", bg="#eef3f4", fg="#29495e", font=("Segoe UI", 9, "bold")).grid(row=5, column=0, sticky="w", pady=4)
        tk.Entry(info, textvariable=self.logo_var, font=("Segoe UI", 10), state="readonly").grid(row=5, column=1, columnspan=2, sticky="ew", padx=(0, 8), pady=4, ipady=4)
        self._button(info, "Escolher logo", self.choose_logo).grid(row=5, column=3, sticky="e", padx=(0, 16), pady=4)

        toolbar = tk.Frame(body, bg="#eef3f4")
        toolbar.pack(fill="x", pady=(0, 12))
        self._button(toolbar, "Adicionar vídeos", self.add_files, primary=True).pack(side="left")
        self._button(toolbar, "Remover", self.remove_selected).pack(side="left", padx=8)
        self._button(toolbar, "Subir", lambda: self.move(-1)).pack(side="left")
        self._button(toolbar, "Descer", lambda: self.move(1)).pack(side="left", padx=8)

        columns = ("ordem", "arquivo", "inicio", "duracao")
        self.tree = ttk.Treeview(body, columns=columns, show="headings", height=7)
        self.tree.heading("ordem", text="#")
        self.tree.heading("arquivo", text="Vídeo")
        self.tree.heading("inicio", text="Início identificado")
        self.tree.heading("duracao", text="Duração")
        self.tree.column("ordem", width=55, anchor="center", stretch=False)
        self.tree.column("arquivo", width=520)
        self.tree.column("inicio", width=165, anchor="center", stretch=False)
        self.tree.column("duracao", width=110, anchor="center", stretch=False)
        self.tree.pack(fill="both", expand=True)

        destination = tk.Frame(body, bg="#eef3f4")
        destination.pack(fill="x", pady=14)
        tk.Label(destination, text="Salvar reunião em:", bg="#eef3f4", fg="#19384d", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        row = tk.Frame(destination, bg="#eef3f4")
        row.pack(fill="x", pady=(5, 0))
        tk.Entry(row, textvariable=self.output_var, font=("Segoe UI", 10)).pack(side="left", fill="x", expand=True, ipady=6)
        self._button(row, "Escolher pasta", self.choose_output).pack(side="left", padx=(8, 0))

        progress_row = tk.Frame(body, bg="#eef3f4")
        progress_row.pack(fill="x", pady=(2, 8))
        self.progress = ttk.Progressbar(progress_row, variable=self.progress_var, maximum=100)
        self.progress.pack(side="left", fill="x", expand=True)
        tk.Label(progress_row, textvariable=self.progress_text_var, width=7, bg="#eef3f4", fg="#19384d", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(10, 0))
        tk.Label(body, textvariable=self.status_var, bg="#eef3f4", fg="#29495e", font=("Segoe UI", 10)).pack(anchor="w")
        self.process_button = self._button(body, "Processar reunião completa", self.start_processing, primary=True)
        self.process_button.pack(anchor="e", pady=(12, 0))

    def _button(self, parent: tk.Widget, text: str, command, primary: bool = False) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg="#245371" if primary else "white",
            fg="white" if primary else "#245371",
            activebackground="#3f7e99",
            activeforeground="white",
            relief="flat",
            padx=14,
            pady=8,
            font=("Segoe UI", 10, "bold"),
            cursor="hand2",
        )

    def add_files(self) -> None:
        names = filedialog.askopenfilenames(
            title="Selecione os vídeos na sequência da reunião",
            filetypes=[("Vídeos", "*.mp4 *.mkv *.avi *.mov"), ("Todos os arquivos", "*.*")],
        )
        existing = {str(item).lower() for item in self.files}
        self.files.extend(pathlib.Path(name) for name in names if name.lower() not in existing)
        self.files.sort(key=natural_key)
        if self.files:
            first_start = recording_start(self.files[0])
            if first_start:
                self.date_var.set(first_start.strftime("%d/%m/%Y"))
                if not self.time_var.get().strip():
                    self.time_var.set(first_start.strftime("%H:%M:%S"))
        self.refresh()

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, path in enumerate(self.files, 1):
            self.tree.insert(
                "", "end", iid=str(index - 1),
                values=(index, path.name, format_recording_time(recording_start(path)), "a calcular"),
            )
        self.status_var.set(f"{len(self.files)} vídeo(s) na fila. Confira a ordem antes de processar.")

    def remove_selected(self) -> None:
        indexes = sorted((int(value) for value in self.tree.selection()), reverse=True)
        for index in indexes:
            self.files.pop(index)
        self.refresh()

    def move(self, direction: int) -> None:
        selection = self.tree.selection()
        if len(selection) != 1:
            return
        index = int(selection[0])
        target = index + direction
        if 0 <= target < len(self.files):
            self.files[index], self.files[target] = self.files[target], self.files[index]
            self.refresh()
            self.tree.selection_set(str(target))

    def choose_output(self) -> None:
        selected = filedialog.askdirectory(title="Escolha onde salvar as reuniões")
        if selected:
            self.output_var.set(selected)

    def choose_logo(self) -> None:
        selected = filedialog.askopenfilename(
            title="Escolha a logomarca do cliente",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg"), ("Todos os arquivos", "*.*")],
        )
        if selected:
            self.logo_var.set(selected)

    def start_processing(self) -> None:
        if not self.files:
            messagebox.showwarning("Nenhum vídeo", "Adicione pelo menos um vídeo antes de processar.")
            return
        self.process_button.configure(state="disabled")
        self.started_at = time.monotonic()
        self._tick_elapsed()
        threading.Thread(target=self.process, daemon=True).start()

    def update_status(self, text: str, progress: float) -> None:
        self.current_status = text
        self.current_progress = max(0, min(100, progress))
        self.after(0, self._render_status)

    def _render_status(self) -> None:
        elapsed = format_duration(time.monotonic() - self.started_at) if self.started_at else "00:00:00"
        self.status_var.set(f"{self.current_status}  •  Tempo decorrido: {elapsed}")
        self.progress_var.set(self.current_progress)
        self.progress_text_var.set(f"{self.current_progress:.0f}%")

    def _tick_elapsed(self) -> None:
        if self.started_at is None:
            return
        self._render_status()
        self.after(1000, self._tick_elapsed)

    def process(self) -> None:
        try:
            stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            output = pathlib.Path(self.output_var.get()).expanduser() / f"reuniao_{stamp}"
            audio_parts = output / "audios_por_video"
            audio_parts.mkdir(parents=True, exist_ok=False)

            durations: list[float] = []
            mp3_files: list[pathlib.Path] = []
            for index, video in enumerate(self.files, 1):
                self.update_status(f"Extraindo áudio {index}/{len(self.files)}: {video.name}", index / (len(self.files) + 2) * 35)
                durations.append(duration_seconds(video))
                target = audio_parts / f"{index:02d}_{safe_name(video.stem)}.mp3"
                run_ffmpeg(["-i", str(video), "-vn", "-ac", "1", "-ar", "16000", "-b:a", "96k", str(target), "-y"])
                mp3_files.append(target)

            list_file = output / "lista_audios.txt"
            list_file.write_text(
                "\n".join("file '" + str(path.resolve()).replace("'", "'\\''") + "'" for path in mp3_files),
                encoding="utf-8",
            )
            combined = output / "AUDIO_REUNIAO_COMPLETA.mp3"
            self.update_status("Unindo todos os áudios da reunião...", 40)
            run_ffmpeg(["-f", "concat", "-safe", "0", "-i", str(list_file), "-c:a", "libmp3lame", "-b:a", "96k", str(combined), "-y"])

            self.update_status("Preparando o modelo de transcrição...", 48)
            def transcription_progress(position: float, total: float) -> None:
                ratio = position / total if total else 0
                self.update_status(
                    f"Transcrevendo {format_duration(position)} de {format_duration(total)}...",
                    48 + ratio * 42,
                )
            raw_rows, language, probability = transcribe(combined, sum(durations), transcription_progress)
            offsets = []
            accumulated = 0.0
            for duration in durations:
                offsets.append(accumulated)
                accumulated += duration
            rows = []
            for row in raw_rows:
                origin_index = max(index for index, offset in enumerate(offsets) if offset <= row["inicio"])
                from datetime import timedelta

                relative_start = row["inicio"] - offsets[origin_index]
                relative_end = row["fim"] - offsets[origin_index]
                source_start = recording_start(self.files[origin_index])
                if source_start:
                    spoken_start = source_start + timedelta(seconds=max(0, relative_start))
                    spoken_end = source_start + timedelta(seconds=max(0, relative_end))
                    start_label = spoken_start.strftime("%d/%m/%Y %H:%M:%S")
                    end_label = spoken_end.strftime("%H:%M:%S")
                else:
                    start_label = "horário não identificado"
                    end_label = ""
                rows.append({
                    **row,
                    "arquivo": self.files[origin_index].name,
                    "data_hora_inicio": start_label,
                    "data_hora_fim": end_label,
                })

            metadata = {
                "idioma": language,
                "probabilidade_idioma": probability,
                "duracao_total": sum(durations),
                "criado_em": datetime.now().isoformat(timespec="seconds"),
            }
            meeting = {
                "cliente": self.client_var.get().strip(),
                "titulo": self.title_var.get().strip(),
                "data": self.date_var.get().strip(),
                "horario": self.time_var.get().strip(),
                "local": self.location_var.get().strip(),
                "participantes": self.participants_text.get("1.0", "end").strip(),
                "observacoes": self.notes_text.get("1.0", "end").strip(),
                "logo": self.logo_var.get().strip(),
            }
            self.update_status("Montando os relatórios em Word...", 94)
            write_outputs(output, self.files, rows, metadata, meeting)
            list_file.unlink(missing_ok=True)
            self.update_status(f"Concluído. Áudios e relatórios salvos em {output}", 100)
            self.after(0, lambda: messagebox.showinfo("Reunião processada", f"Áudio, transcrição e ATA foram salvos em:\n\n{output}"))
        except Exception as exc:
            self.update_status(f"Não foi possível concluir: {exc}", 0)
            self.after(0, lambda error=str(exc): messagebox.showerror("Erro no processamento", error))
        finally:
            self.after(0, lambda: self.process_button.configure(state="normal"))
            self.started_at = None


if __name__ == "__main__":
    AtaApp().mainloop()
