"""Atlas Copco email signature generator."""

from __future__ import annotations

import argparse
import html
import re
import sys
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path


BELGIAN_NATIONAL_PATTERN = re.compile(r"^04\d{8}$")
BELGIAN_INTERNATIONAL_PATTERN = re.compile(r"^\+32\d{9}$")
GERMAN_PATTERN = re.compile(r"^\+49(?:\(0\))?\s?\d{7,13}$")
EMAIL_DOMAIN = "@atlascopco.com"


@dataclass
class SignatureData:
    name: str
    title: str
    email: str
    mobile: str
    phone: str | None = None


class ValidationError(ValueError):
    """Raised when the supplied data cannot produce a signature."""


def sanitize_tel(value: str) -> str:
    """Return a sanitised phone number for use in tel: hyperlinks."""

    sanitized = re.sub(r"[\s()-]", "", value)
    sanitized = re.sub(r"^\+320", "+32", sanitized)
    sanitized = re.sub(r"^\+490", "+49", sanitized)
    return re.sub(r"[^+0-9]", "", sanitized)


def sanitize_file_segment(value: str) -> str:
    """Return a filesystem-safe segment derived from *value*."""

    return re.sub(r"[^a-zA-Z0-9@._+-]", "_", value)


def escape_rtf(value: str) -> str:
    """Escape text for embedding in an RTF document."""

    return (
        value.replace("\\", "\\\\")
        .replace("{", "\\{")
        .replace("}", "\\}")
    )


def is_valid_mobile_number(number: str) -> bool:
    """Validate Belgian and German phone number formats."""

    trimmed = number.strip()
    clean = re.sub(r"[\s()-]", "", number)
    normalized = re.sub(r"^\+320", "+32", clean)

    if BELGIAN_NATIONAL_PATTERN.fullmatch(normalized):
        return True

    if BELGIAN_INTERNATIONAL_PATTERN.fullmatch(normalized):
        return True

    german_candidate = re.sub(r"\s+", " ", trimmed)
    return GERMAN_PATTERN.fullmatch(german_candidate) is not None


def validate_inputs(data: SignatureData) -> None:
    """Raise :class:`ValidationError` if the supplied inputs are invalid."""

    errors: list[str] = []

    if not data.name.strip():
        errors.append("Name is required.")

    if not data.title.strip():
        errors.append("Title is required.")

    email = data.email.strip()
    if not email:
        errors.append("Email address is required.")
    elif not email.lower().endswith(EMAIL_DOMAIN):
        errors.append(f"Email address must end with '{EMAIL_DOMAIN}'.")

    mobile = data.mobile.strip()
    if not mobile:
        errors.append("Mobile number is required.")
    elif not is_valid_mobile_number(mobile):
        errors.append(
            "Mobile number must be a valid Belgian (04xxxxxxxx or +32xxxxxxxxx) "
            "or German (+49(0) 123456789) number."
        )

    if errors:
        raise ValidationError("\n".join(errors))


def build_signatures(data: SignatureData) -> tuple[str, str, str]:
    """Return the HTML, plain text and RTF signatures."""

    validate_inputs(data)

    safe_name = html.escape(data.name.strip())
    safe_title = html.escape(data.title.strip())
    safe_email = html.escape(data.email.strip())
    mobile = data.mobile.strip()
    phone = (data.phone or "").strip()
    safe_mobile = html.escape(mobile)
    safe_phone = html.escape(phone)

    tel_mobile = sanitize_tel(mobile)
    tel_phone = sanitize_tel(phone) if phone else ""

    contact_info_html = textwrap.dedent(
        f"""
        <div style="font-size:9pt;color:#2F363A;">Mobile: <a href="tel:{tel_mobile}" style="color:#2F363A;text-decoration:none;">{safe_mobile}</a></div>
        """
    ).strip()

    contact_info_text = f"Mobile: {mobile}\n"

    if phone:
        contact_info_html += textwrap.dedent(
            f"""
            <div style="font-size:9pt;color:#2F363A;">Phone: <a href="tel:{tel_phone}" style="color:#2F363A;text-decoration:none;">{safe_phone}</a></div>
            """
        ).strip()
        contact_info_text += f"Phone: {phone}\n"

    contact_info_html += textwrap.dedent(
        f"""
        <div style="font-size:9pt;color:#2F363A;">E-mail: <a href="mailto:{safe_email}" style="color:#0092BC;text-decoration:none;">{safe_email}</a></div>
        """
    ).strip()
    contact_info_text += f"E-mail: {data.email.strip()}\n"

    html_signature = textwrap.dedent(
        f"""
        <html><head><meta charset="UTF-8"></head><body>
        <div style="font-family:Calibri,Arial,sans-serif; font-size:9pt; color:#2F363A; line-height:1.4; max-width:600px;">
          <div>Best regards,</div><br>
          <div style="font-weight:bold; color:#0092BC; font-size:11pt; margin-top:4px;">{safe_name}</div>
          <div style="font-weight:bold; font-size:9pt; margin-top:2px;">{safe_title}</div>

          <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

          <div style="font-weight:bold; font-size:10pt;">Atlas Copco Tools Central Europe GmbH</div>
          <div>Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen</div>
          <div style="color:red; font-weight:bold;">Head Office: <u>New Address as of Dec 1, 2025</u>: Wetterschacht 9, D-45139 Essen</div>
          <div>Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen</div>
          <div>Managing Director: Claus Schiedeck, Peter Edmonds</div><br>

          {contact_info_html}

          <div style="margin-top:6px;">VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)</div>
          <div>Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)</div>
          <div>Peppol ID (Belgium): 0208:0473470658</div>

          <div style="margin-top:6px;">Privacy Policy: <a href="https://www.atlascopco.com/nl-be/itba/privacy-policy" target="_blank" style="color:#0092BC;text-decoration:none;">https://www.atlascopco.com/nl-be/itba/privacy-policy</a></div>

          <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

          <div><b style="color:#0092BC;">Level up your experience at</b> <a href="https://www.atlascopco.com" target="_blank" style="color:#0092BC; font-style:italic; text-decoration:none;">atlascopco.com</a></div>

          <hr style="border:0;border-top:1px solid #E1D6CE;margin:8px 0;">

          <table cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td><img src="https://groupapp.atlascopco.com/mail-signature/Signature/AtlasCopco/ACBlueSquare.png" width="8" height="8" style="display:block;"></td>
              <td style="padding-left:5px;">
                <span style="font-size:8pt; color:#A1A9B4; font-weight:bold;">Part of Atlas Copco Group</span>
              </td>
            </tr>
          </table>
        </div>
        </body></html>
        """
    ).strip()

    text_signature = textwrap.dedent(
        f"""
        Best regards,
        {data.name.strip()}
        {data.title.strip()}

        ----------------------------------------
        Atlas Copco Tools Central Europe GmbH
        Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen
        Head Office: New Address as of Dec 1, 2025: Wetterschacht 9, D-45139 Essen
        Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen
        Managing Director: Claus Schiedeck, Peter Edmonds

        {contact_info_text}VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)
        Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)
        Peppol ID (Belgium): 0208:0473470658

        Privacy Policy: https://www.atlascopco.com/nl-be/itba/privacy-policy
        ----------------------------------------
        Level up your experience at atlascopco.com
        ----------------------------------------
        Part of Atlas Copco Group
        """
    ).strip() + "\n"

    rtf_contact_info = contact_info_text.replace("\n", "\\line ")
    rtf_signature = textwrap.dedent(
        f"""
        {{\\rtf1\\ansi\\deff0
        {{\\fonttbl{{\\f0 Calibri;}}}}
        \\f0\\fs18 Best regards,\\line
        \\b\\cf1 {escape_rtf(data.name.strip())}\\b0\\line
        {escape_rtf(data.title.strip())}\\line
        \\line
        Atlas Copco Tools Central Europe GmbH\\line
        Head Office: Atlas Copco Tools Central Europe GmbH, Langemarckstr. 35, D-45141 Essen\\line
        Head Office: New Address as of Dec 1, 2025: Wetterschacht 9, D-45139 Essen\\line
        Branch Office Belgium (Atlas Copco Tools Belgium): Bremakker 45, B-3740 Bilzen\\line
        Managing Director: Claus Schiedeck, Peter Edmonds\\line
        \\line
        {escape_rtf(rtf_contact_info)}\\line
        VAT Reg.No.: DE811155641 (Germany) / BE0473470658 (Belgium)\\line
        Company Reg. Office: HRB 5096 – Local Court Essen (Head Office) / R.C.B. 646980 - Brussels (Belgian Branch)\\line
        Peppol ID (Belgium): 0208:0473470658\\line
        \\line
        Privacy Policy: https://www.atlascopco.com/nl-be/itba/privacy-policy\\line
        Level up your experience at atlascopco.com\\line
        Part of Atlas Copco Group
        }}
        """
    ).strip()

    return html_signature, text_signature, rtf_signature


def write_zip(html_signature: str, text_signature: str, rtf_signature: str, *, email: str, output: Path) -> Path:
    """Write the generated signatures to a ZIP archive."""

    base_name = f"AtlasCopco({sanitize_file_segment(email.strip())})"
    archive_path = output / f"{base_name}.zip"

    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{base_name}.htm", html_signature)
        archive.writestr(f"{base_name}.txt", text_signature)
        archive.writestr(f"{base_name}.rtf", rtf_signature)
        archive.writestr(f"{base_name}_files/", "")

    return archive_path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate Atlas Copco email signature files.")
    parser.add_argument("name", help="Full name of the employee")
    parser.add_argument("title", help="Job title")
    parser.add_argument("email", help="Atlas Copco email address")
    parser.add_argument("mobile", help="Primary mobile phone number")
    parser.add_argument("phone", nargs="?", default="", help="Optional landline phone number")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path.cwd(),
        help="Directory where the ZIP archive will be created",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    output_dir: Path = args.output
    output_dir.mkdir(parents=True, exist_ok=True)

    data = SignatureData(
        name=args.name,
        title=args.title,
        email=args.email,
        mobile=args.mobile,
        phone=args.phone or None,
    )

    try:
        html_signature, text_signature, rtf_signature = build_signatures(data)
    except ValidationError as exc:  # pragma: no cover - CLI surface
        print(f"Error:\n{exc}", file=sys.stderr)
        return 1

    archive_path = write_zip(
        html_signature,
        text_signature,
        rtf_signature,
        email=data.email,
        output=output_dir,
    )

    print(f"Signature bundle created: {archive_path}")
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI surface
    raise SystemExit(main())
