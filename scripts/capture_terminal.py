#!/usr/bin/env python3
"""Run a terminal command and render its captured output as a PNG screenshot."""

from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ANSI_RE = re.compile(r"\x1b(?:\[[0-?]*[ -/]*[@-~]|\][^\x07]*(?:\x07|\x1b\\))")
CANVAS_WIDTH = 1180
HEADER_HEIGHT = 48
FOOTER_HEIGHT = 42
PADDING_X = 24
PADDING_Y = 20
FONT_SIZE = 16
LINE_GAP = 7


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\msyh.ttc"),
        Path(r"C:\Windows\Fonts\consola.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            try:
                return ImageFont.truetype(str(candidate), size)
            except OSError:
                continue
    return ImageFont.load_default()


def decode_output(data: bytes) -> str:
    for encoding in ("utf-8", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def run_command(command: str, shell: str, cwd: Path) -> tuple[str, int]:
    if shell == "powershell":
        argv = ["powershell", "-NoProfile", "-Command", command]
    elif shell == "cmd":
        argv = ["cmd.exe", "/d", "/s", "/c", command]
    elif shell == "bash":
        argv = ["bash", "-lc", command]
    elif shell == "wsl":
        argv = ["wsl.exe", "bash", "-lc", command]
    else:
        raise ValueError(f"unsupported shell: {shell}")

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    completed = subprocess.run(
        argv,
        cwd=cwd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        env=env,
        check=False,
    )
    return decode_output(completed.stdout), completed.returncode


def wrap_text(
    text: str,
    draw: ImageDraw.ImageDraw,
    font: ImageFont.ImageFont,
    max_width: float,
) -> list[str]:
    lines: list[str] = []
    for source_line in text.replace("\r", "").split("\n"):
        source_line = source_line.replace("\t", "    ")
        if not source_line:
            lines.append("")
            continue
        current = ""
        for char in source_line:
            candidate = current + char
            if current and draw.textlength(candidate, font=font) > max_width:
                lines.append(current)
                current = char
            else:
                current = candidate
        lines.append(current)
    return lines


def trim_lines(lines: list[str], limit: int) -> list[str]:
    if len(lines) <= limit:
        return lines
    keep_head = 4
    keep_tail = max(1, limit - keep_head - 1)
    omitted = len(lines) - keep_head - keep_tail
    return (
        lines[:keep_head]
        + [f"... {omitted} lines omitted ..."]
        + lines[-keep_tail:]
    )


def apply_redactions(text: str, patterns: list[str]) -> str:
    for pattern in patterns:
        text = re.sub(pattern, "[REDACTED]", text)
    return text


def render_capture(
    output: Path,
    title: str,
    command: str,
    output_text: str,
    return_code: int,
    max_lines: int,
) -> None:
    clean_output = ANSI_RE.sub("", output_text)
    font = load_font(FONT_SIZE)
    small = load_font(13)
    draw_probe = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    content_width = CANVAS_WIDTH - 2 * PADDING_X
    body_lines = trim_lines(
        wrap_text(clean_output.rstrip() or "<no output>", draw_probe, font, content_width),
        max_lines,
    )
    command_lines = wrap_text(command, draw_probe, font, content_width)
    line_height = FONT_SIZE + LINE_GAP
    body_height = PADDING_Y * 2 + line_height * (len(command_lines) + 1 + len(body_lines))
    total_height = HEADER_HEIGHT + body_height + FOOTER_HEIGHT

    image = Image.new("RGB", (CANVAS_WIDTH, total_height), "#0f172a")
    draw = ImageDraw.Draw(image)

    draw.rectangle((0, 0, CANVAS_WIDTH, HEADER_HEIGHT), fill="#1e293b")
    for index, color in enumerate(("#ef4444", "#f59e0b", "#22c55e")):
        cx = 24 + index * 22
        draw.ellipse((cx - 6, 18, cx + 6, 30), fill=color)
    draw.text((112, 14), title, fill="#e2e8f0", font=small)

    y = HEADER_HEIGHT + PADDING_Y
    for line in command_lines:
        draw.text((PADDING_X, y), line, fill="#86efac", font=font)
        y += line_height
    y += LINE_GAP
    for line in body_lines:
        draw.text((PADDING_X, y), line, fill="#e5e7eb", font=font)
        y += line_height

    footer_top = total_height - FOOTER_HEIGHT
    draw.rectangle((0, footer_top, CANVAS_WIDTH, total_height), fill="#1e293b")
    status = "EXIT 0" if return_code == 0 else f"EXIT {return_code}"
    status_color = "#86efac" if return_code == 0 else "#fca5a5"
    draw.text((PADDING_X, footer_top + 12), status, fill=status_color, font=small)
    timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    draw.text((CANVAS_WIDTH - 330, footer_top + 12), timestamp, fill="#94a3b8", font=small)

    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--command", help="command to execute and capture")
    source.add_argument("--input", type=Path, help="existing UTF-8 text to render")
    parser.add_argument("-o", "--output", type=Path, required=True)
    parser.add_argument("--title", default="Terminal capture")
    parser.add_argument(
        "--shell",
        choices=("powershell", "cmd", "bash", "wsl"),
        default="powershell",
    )
    parser.add_argument("--cwd", type=Path, default=Path.cwd())
    parser.add_argument("--max-lines", type=int, default=44)
    parser.add_argument("--allow-failure", action="store_true")
    parser.add_argument(
        "--redact",
        action="append",
        default=[],
        help="regular expression to replace with [REDACTED]; may be repeated",
    )
    parser.add_argument(
        "--transcript",
        action="store_true",
        help="also write a .txt transcript beside the PNG",
    )
    args = parser.parse_args()

    if args.command:
        output_text, return_code = run_command(
            args.command,
            args.shell,
            args.cwd.resolve(),
        )
        command = apply_redactions(args.command, args.redact)
    else:
        command = f"cat {args.input}"
        output_text = args.input.read_text(encoding="utf-8", errors="replace")
        return_code = 0
    output_text = apply_redactions(output_text, args.redact)

    render_capture(
        args.output.resolve(),
        args.title,
        command,
        output_text,
        return_code,
        args.max_lines,
    )
    if args.transcript:
        transcript = args.output.with_suffix(".txt")
        transcript.write_text(
            f"$ {command}\n\n{output_text}\n\n[{return_code}]\n",
            encoding="utf-8",
        )
    print(f"captured: {args.output.resolve()}")
    return 0 if return_code == 0 or args.allow_failure else return_code


if __name__ == "__main__":
    raise SystemExit(main())
