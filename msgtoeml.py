#!/usr/bin/env python3
import argparse
import sys
from email.message import EmailMessage
from email.utils import formatdate

import extract_msg


def to_text(value, fallback_charset="utf-8"):
    """
    Normalize extract-msg fields that might be str/bytes/None into a clean str.
    """
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, bytes):
        # Try UTF-8 first; fall back to latin-1 as a last resort
        try:
            return value.decode(fallback_charset, errors="replace")
        except Exception:
            return value.decode("latin-1", errors="replace")
    # Anything else -> string representation
    return str(value)


def build_eml_from_msg(m: extract_msg.Message) -> EmailMessage:
    eml = EmailMessage()

    # Headers
    eml["Subject"] = to_text(getattr(m, "subject", ""))
    eml["From"] = to_text(getattr(m, "sender", ""))
    eml["To"] = to_text(getattr(m, "to", ""))

    cc = to_text(getattr(m, "cc", ""))
    if cc:
        eml["Cc"] = cc

    # Date
    msg_date = getattr(m, "date", None)
    if msg_date:
        eml["Date"] = to_text(msg_date)
    else:
        eml["Date"] = formatdate(localtime=True)

    # Body (plain text)
    text_body = to_text(getattr(m, "body", ""))
    # Force charset so clients render consistently
    eml.set_content(text_body, subtype="plain", charset="utf-8")

    # HTML alternative (if present)
    html_body = getattr(m, "htmlBody", None)
    if html_body is None:
        html_body = getattr(m, "htmlbody", None)

    html_body = to_text(html_body)
    if html_body.strip():
        eml.add_alternative(html_body, subtype="html", charset="utf-8")

    return eml


def main():
    parser = argparse.ArgumentParser(description="Convert Outlook .msg to .eml")
    parser.add_argument("input_msg", help="Path to input .msg file")
    parser.add_argument("output_eml", help="Path to output .eml file")
    args = parser.parse_args()

    m = None
    try:
        m = extract_msg.Message(args.input_msg)
        eml = build_eml_from_msg(m)
    except Exception as e:
        print(f"[-] Failed to parse MSG: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        try:
            if m:
                m.close()
        except Exception:
            pass

    with open(args.output_eml, "wb") as f:
        f.write(eml.as_bytes())

    print(f"[+] Wrote EML: {args.output_eml}")


if __name__ == "__main__":
    main()
