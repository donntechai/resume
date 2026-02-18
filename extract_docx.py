import sys


def main() -> int:
    try:
        from docx import Document  # type: ignore
    except Exception:
        print(
            "Missing dependency: python-docx. Install with:\n"
            "  python -m pip install python-docx",
            file=sys.stderr,
        )
        return 2

    if len(sys.argv) != 2:
        print("Usage: python extract_docx.py <path-to-docx>", file=sys.stderr)
        return 2

    docx_path = sys.argv[1]
    doc = Document(docx_path)

    lines: list[str] = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            lines.append("")
            continue
        lines.append(t)

    # Normalize excessive blank lines.
    out: list[str] = []
    blank_run = 0
    for line in lines:
        if line == "":
            blank_run += 1
            if blank_run <= 1:
                out.append("")
        else:
            blank_run = 0
            out.append(line)

    print("\n".join(out).strip() + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
