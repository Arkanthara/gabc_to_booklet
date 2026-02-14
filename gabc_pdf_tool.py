import os
import re
import sys
from gooey import Gooey, GooeyParser
from pypdf import PdfReader, PdfWriter

# ----------------------------------------------------------------------
# GABC Parser functions (copied from gabc_parser.py)
# ----------------------------------------------------------------------
def split_by_empty_lines(content: str) -> list:
    """Split content into entries separated by one or more empty/whitespace-only lines."""
    blocks = re.split(r"\n\s*\n", content)
    return [block.strip() for block in blocks if block.strip()]

def normalize_text(text: str) -> str:
    """Strip punctuation, lowercase, replace spaces with underscores."""
    cleaned = re.sub(r"[^\w\s\-]", "", text)
    return cleaned.strip().lower().replace(" ", "_")

def extract_field(entry: str, field: str) -> str:
    """Extract value of a given field (e.g., 'name', 'office-part')."""
    match = re.search(rf"^{field}:\s*(.+?);?\s*$", entry, re.MULTILINE | re.IGNORECASE)
    return normalize_text(match.group(1)) if match else ""

def is_significant(entry: str) -> bool:
    """Check if entry has at least one of 'name:' or 'office-part:'."""
    return bool(extract_field(entry, "name") or extract_field(entry, "office-part"))

def transform_annotation(entry: str) -> str:
    """
    For every line starting with 'annotation:' (case‑insensitive),
    wrap its content in \textsc{...}. Preserves any trailing semicolon.
    """
    lines = entry.splitlines()
    new_lines = []
    for line in lines:
        m = re.match(r'^(annotation:)(\s*)(.+?)(;?)$', line, re.IGNORECASE)
        if m:
            prefix, spaces, content, semicolon = m.groups()
            new_content = "{\\textsc{" + f"{content}" + "}" + "}"
            new_line = f"{prefix}{spaces}{new_content}{semicolon}"
            new_lines.append(new_line)
        else:
            new_lines.append(line)
    return "\n".join(new_lines)

def save_entries_separately(entries: list, output_dir: str, create_dir: bool = True) -> int:
    """Save each *significant* entry to its own .gabc file, after transforming annotation fields."""
    if create_dir:
        os.makedirs(output_dir, exist_ok=True)
    elif not os.path.isdir(output_dir):
        raise ValueError(f"Output directory '{output_dir}' does not exist.")

    used_names = {}
    written = 0

    for i, entry in enumerate(entries):
        if not is_significant(entry):
            continue

        entry = transform_annotation(entry)

        office = extract_field(entry, "office-part")
        name = extract_field(entry, "name")

        base = office or name or f"unknown_{i + 1}"

        count = used_names.get(base, 0) + 1
        used_names[base] = count
        filename = f"{base}.gabc" if count == 1 else f"{base}_{count}.gabc"

        path = os.path.join(output_dir, filename)

        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(entry.rstrip() + "\n")
            written += 1
        except Exception as e:
            print(f"⚠️  Failed to write '{path}': {e}")

    return written

# ----------------------------------------------------------------------
# PDF Booklet functions (copied from pdf_booklet.py)
# ----------------------------------------------------------------------
def generate_booklet_pages(nop):
    """
    Generate the booklet page order for a document with `nop` pages.
    Returns a list of page numbers (1‑based) in the required order.
    """
    nop_booklet = (nop + 3) // 4          # ceiling division
    base = [2 * (i + 1) for i in range(nop_booklet)][::-1]
    num = nop_booklet * 4 + 1
    pages = []
    for i in base:
        pages.append(i)
        pages.append(num - i)
        pages.append(num - i + 1)
        pages.append(num - (num - i + 1))
    return pages

def booklet_rearrange(input_path, start, end, print_only, output_path=None):
    """
    Core logic: read PDF, rearrange pages, and write the new PDF or print the order.
    Blank pages are automatically inserted when needed.
    """
    try:
        reader = PdfReader(input_path)
    except Exception as e:
        print(f"Error reading PDF file: {e}")
        sys.exit(1)

    total_pages = len(reader.pages)

    if end is None or end > total_pages:
        end = total_pages
    if start < 1:
        start = 1
    if start > end:
        print("Error: Start page cannot be greater than end page.")
        sys.exit(1)

    nop = end - start + 1
    pages_order = generate_booklet_pages(nop)

    if print_only:
        print("New page order (1‑based, relative to the selected range):")
        print(pages_order)
        return

    if output_path and os.path.exists(output_path):
        print(f"Error: Output file already exists:\n{output_path}\nFile will not be overwritten.")
        sys.exit(1)

    writer = PdfWriter()
    for p in pages_order:
        if p > nop:
            writer.add_blank_page()
        else:
            idx = p + start - 2
            writer.add_page(reader.pages[idx])

    try:
        with open(output_path, 'wb') as f:
            writer.write(f)
        print(f"Success! Rearranged PDF saved to:\n{output_path}")
    except Exception as e:
        print(f"Error writing output file: {e}")
        sys.exit(1)

# ----------------------------------------------------------------------
# Main Gooey application
# ----------------------------------------------------------------------
@Gooey(
    program_name="GABC Parser & PDF Booklet Rearranger",
    program_description="Choose a tool from the dropdown and fill in its fields.",
    clear_before_run=True,
    default_size=(750, 800)
)
def main():
    parser = GooeyParser(description="Select the tool you want to use.")

    # Mode selection
    parser.add_argument(
        'mode',
        metavar='Select Tool',
        choices=['GABC Parser', 'PDF Booklet Rearranger'],
        default='GABC Parser',
        help='Choose which utility to run',
        widget='Dropdown'
    )

    # ------------------- GABC Parser group -------------------
    gabc_group = parser.add_argument_group('GABC Parser Options', 'Fields for splitting a .gabc file')
    gabc_group.add_argument(
        '--gabc_input',
        metavar='GABC Input File',
        help='The .gabc file to split',
        widget='FileChooser'
    )
    gabc_group.add_argument(
        '--gabc_output_dir',
        metavar='GABC Output Folder',
        help='Directory to save the individual .gabc files',
        widget='DirChooser'
    )
    gabc_group.add_argument(
        '--gabc_create_dir',
        metavar='Create output directory if missing',
        action='store_true',
        default=True,
        help='Create the output folder if it does not exist',
        widget='CheckBox',
        gooey_options={'initial_value': True}
    )

    # ------------------- PDF Booklet group -------------------
    pdf_group = parser.add_argument_group('PDF Booklet Options', 'Fields for rearranging a PDF for booklet printing')
    pdf_group.add_argument(
        '--pdf_input',
        metavar='PDF Input File',
        help='The PDF file to rearrange',
        widget='FileChooser'
    )
    pdf_group.add_argument(
        '--pdf_output_folder',
        metavar='PDF Output Folder',
        help='Folder where the rearranged PDF will be saved',
        widget='DirChooser'
    )
    pdf_group.add_argument(
        '--pdf_output_filename',
        metavar='PDF Output Filename',
        help='Name of the output PDF file (e.g., "mybooklet.pdf")',
        widget='TextField'
    )
    pdf_group.add_argument(
        '--pdf_start',
        metavar='Start Page',
        type=int,
        default=1,
        help='Starting page number (default: 1)'
    )
    pdf_group.add_argument(
        '--pdf_end',
        metavar='End Page',
        type=int,
        help='Ending page number (default: last page)'
    )
    pdf_group.add_argument(
        '--pdf_print_pages',
        metavar='Only print page order',
        action='store_true',
        help='Only print the new page order, do not create PDF'
    )

    args = parser.parse_args()

    # ----------------------------------------------------------
    # Execute the chosen tool
    # ----------------------------------------------------------
    if args.mode == 'GABC Parser':
        # Validate required fields
        if not args.gabc_input:
            print("❌ Please select an input file for the GABC Parser.")
            sys.exit(1)
        if not args.gabc_output_dir:
            print("❌ Please select an output folder for the GABC Parser.")
            sys.exit(1)

        # Read input file
        try:
            with open(args.gabc_input, "r", encoding="utf-8") as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                with open(args.gabc_input, "r", encoding="latin-1") as f:
                    content = f.read()
            except Exception as e:
                print(f"❌ Could not read file: {e}")
                sys.exit(1)

        entries = split_by_empty_lines(content)
        entries = [e for e in entries if e]   # remove empty blocks

        if not entries:
            print("❌ No entries found in the GABC file.")
            sys.exit(1)

        count = save_entries_separately(entries, args.gabc_output_dir, args.gabc_create_dir)
        print(f"\n✅ Saved {count} significant file(s) to: {os.path.abspath(args.gabc_output_dir)}")

    else:   # PDF Booklet Rearranger
        # Validate required fields
        if not args.pdf_input:
            print("❌ Please select an input PDF file.")
            sys.exit(1)
        if not args.pdf_print_pages:
            if not args.pdf_output_folder:
                print("❌ Please select an output folder for the PDF.")
                sys.exit(1)
            if not args.pdf_output_filename:
                print("❌ Please enter an output filename.")
                sys.exit(1)

        # Prepare output path if needed
        output_full_path = None
        if not args.pdf_print_pages:
            if not args.pdf_output_filename.lower().endswith('.pdf'):
                args.pdf_output_filename += '.pdf'
            output_full_path = os.path.join(args.pdf_output_folder, args.pdf_output_filename)

        # Call the rearrangement function
        booklet_rearrange(
            input_path=args.pdf_input,
            start=args.pdf_start,
            end=args.pdf_end,
            print_only=args.pdf_print_pages,
            output_path=output_full_path
        )

if __name__ == "__main__":
    main()
