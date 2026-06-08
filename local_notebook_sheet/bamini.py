"""Unicode Tamil ↔ Bamini legacy font encoding converter.

Bamini is a legacy Tamil font that maps Tamil glyphs to ASCII/Latin character
positions. This module converts Unicode Tamil text (U+0B80–U+0BFF) to the
Bamini encoding so it displays correctly when rendered with the Bamini font.

The mapping is built by reversing the well-known Bamini dictionary, handling
compound characters (uyir-mei) before single characters to avoid partial
replacements.
"""

from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Unicode → Bamini mapping
# ---------------------------------------------------------------------------
# Built by reversing the standard Bamini encoding table.  Compound characters
# (multi-codepoint sequences like கா, கி, கு …) MUST be listed before their
# base consonants so that the longest match is replaced first.
#
# The dictionary is ordered: longest Unicode sequences first.

_UNICODE_TO_BAMINI: list[tuple[str, str]] = [
    # ── Grantha / special ──
    ("ஸ்ரீ", "="),
    ("ௐ", "≈"),

    # ── க-வரிசை (Ka) ──
    ("கௌ", "nfs"),
    ("கோ", "Nfh"),
    ("கொ", "nfh"),
    ("கா", "fh"),
    ("கி", "fp"),
    ("கீ", "fP"),
    ("கு", "F"),
    ("கூ", "$"),
    ("கெ", "nf"),
    ("கே", "Nf"),
    ("கை", "if"),
    ("க்", "f;"),
    ("க", "f"),

    # ── ங-வரிசை (Nga) ──
    ("ஙௌ", "nqs"),
    ("ஙோ", "Nqh"),
    ("ஙொ", "nqh"),
    ("ஙா", "qh"),
    ("ஙி", "qp"),
    ("ஙீ", "qP"),
    ("ஙெ", "nq"),
    ("ஙே", "Nq"),
    ("ஙை", "iq"),
    ("ங்", "q;"),
    ("ங", "q"),

    # ── ச-வரிசை (Cha) ──
    ("சௌ", "nrs"),
    ("சோ", "Nrh"),
    ("சொ", "nrh"),
    ("சா", "rh"),
    ("சி", "rp"),
    ("சீ", "rP"),
    ("சு", "R"),
    ("சூ", "#"),
    ("செ", "nr"),
    ("சே", "Nr"),
    ("சை", "ir"),
    ("ச்", "r;"),
    ("ச", "r"),

    # ── ஜ-வரிசை (Ja) ──
    ("ஜௌ", "n[s"),
    ("ஜோ", "N[h"),
    ("ஜொ", "n[h"),
    ("ஜா", "[h"),
    ("ஜி", "[p"),
    ("ஜீ", "[P"),
    ("ஜு", "[{"),
    ("ஜூ", "[_"),
    ("ஜெ", "n["),
    ("ஜே", "N["),
    ("ஜை", "i["),
    ("ஜ்", "[;"),
    ("ஜ", "["),

    # ── ஞ-வரிசை (Nya) ──
    ("ஞௌ", "nQs"),
    ("ஞோ", "NQh"),
    ("ஞொ", "nQh"),
    ("ஞா", "Qh"),
    ("ஞி", "Qp"),
    ("ஞீ", "QP"),
    ("ஞெ", "nQ"),
    ("ஞே", "NQ"),
    ("ஞை", "iQ"),
    ("ஞ்", "Q;"),
    ("ஞ", "Q"),

    # ── ட-வரிசை (Ta) ──
    ("டௌ", "nls"),
    ("டோ", "Nlh"),
    ("டொ", "nlh"),
    ("டா", "lh"),
    ("டி", "b"),
    ("டீ", "B"),
    ("டு", "L"),
    ("டூ", "^"),
    ("டெ", "nl"),
    ("டே", "Nl"),
    ("டை", "il"),
    ("ட்", "l;"),
    ("ட", "l"),

    # ── ண-வரிசை (Na - retroflex) ──
    ("ணௌ", "nzs"),
    ("ணோ", "Nzh"),
    ("ணொ", "nzh"),
    ("ணா", "zh"),
    ("ணி", "zp"),
    ("ணீ", "zP"),
    ("ணு", "Z"),
    ("ணூ", "Zh"),
    ("ணெ", "nz"),
    ("ணே", "Nz"),
    ("ணை", "iz"),
    ("ண்", "z;"),
    ("ண", "z"),

    # ── த-வரிசை (Tha) ──
    ("தௌ", "njs"),
    ("தோ", "Njh"),
    ("தொ", "njh"),
    ("தா", "jh"),
    ("தி", "jp"),
    ("தீ", "jP"),
    ("து", "J"),
    ("தூ", "Jh"),
    ("தெ", "nj"),
    ("தே", "Nj"),
    ("தை", "ij"),
    ("த்", "j;"),
    ("த", "j"),

    # ── ந-வரிசை (Na - dental) ──
    ("நௌ", "nes"),
    ("நோ", "Neh"),
    ("நொ", "neh"),
    ("நா", "eh"),
    ("நி", "ep"),
    ("நீ", "eP"),
    ("நு", "E"),
    ("நூ", "Eh"),
    ("நெ", "ne"),
    ("நே", "Ne"),
    ("நை", "ie"),
    ("ந்", "e;"),
    ("ந", "e"),

    # ── ன-வரிசை (Na - alveolar) ──
    ("னௌ", "nds"),
    ("னோ", "Ndh"),
    ("னொ", "ndh"),
    ("னா", "dh"),
    ("னி", "dp"),
    ("னீ", "dP"),
    ("னு", "D"),
    ("னூ", "Dh"),
    ("னெ", "nd"),
    ("னே", "Nd"),
    ("னை", "id"),
    ("ன்", "d;"),
    ("ன", "d"),

    # ── ப-வரிசை (Pa) ──
    ("பௌ", "ngs"),
    ("போ", "Ngh"),
    ("பொ", "ngh"),
    ("பா", "gh"),
    ("பி", "gp"),
    ("பீ", "gP"),
    ("பு", "G"),
    ("பூ", "G+"),
    ("பெ", "ng"),
    ("பே", "Ng"),
    ("பை", "ig"),
    ("ப்", "g;"),
    ("ப", "g"),

    # ── ம-வரிசை (Ma) ──
    ("மௌ", "nks"),
    ("மோ", "Nkh"),
    ("மொ", "nkh"),
    ("மா", "kh"),
    ("மி", "kp"),
    ("மீ", "kP"),
    ("மு", "K"),
    ("மூ", "%"),
    ("மெ", "nk"),
    ("மே", "Nk"),
    ("மை", "ik"),
    ("ம்", "k;"),
    ("ம", "k"),

    # ── ய-வரிசை (Ya) ──
    ("யௌ", "nas"),
    ("யோ", "Nah"),
    ("யொ", "nah"),
    ("யா", "ah"),
    ("யி", "ap"),
    ("யீ", "aP"),
    ("யு", "A"),
    ("யூ", "A+"),
    ("யெ", "na"),
    ("யே", "Na"),
    ("யை", "ia"),
    ("ய்", "a;"),
    ("ய", "a"),

    # ── ர-வரிசை (Ra) ──
    ("ரௌ", "nus"),
    ("ரோ", "Nuh"),
    ("ரொ", "nuh"),
    ("ரா", "uh"),
    ("ரி", "up"),
    ("ரீ", "uP"),
    ("ரு", "U"),
    ("ரூ", "&"),
    ("ரெ", "nu"),
    ("ரே", "Nu"),
    ("ரை", "iu"),
    ("ர்", "u;"),
    ("ர", "u"),

    # ── ல-வரிசை (La) ──
    ("லௌ", "nys"),
    ("லோ", "Nyh"),
    ("லொ", "nyh"),
    ("லா", "yh"),
    ("லி", "yp"),
    ("லீ", "yP"),
    ("லு", "Y"),
    ("லூ", "Yh"),
    ("லெ", "ny"),
    ("லே", "Ny"),
    ("லை", "iy"),
    ("ல்", "y;"),
    ("ல", "y"),

    # ── ள-வரிசை (Lla) ──
    ("ளௌ", "nss"),
    ("ளோ", "Nsh"),
    ("ளொ", "nsh"),
    ("ளா", "sh"),
    ("ளி", "sp"),
    ("ளீ", "sP"),
    ("ளு", "S"),
    ("ளூ", "Sh"),
    ("ளெ", "ns"),
    ("ளே", "Ns"),
    ("ளை", "is"),
    ("ள்", "s;"),
    ("ள", "s"),

    # ── வ-வரிசை (Va) ──
    ("வௌ", "ntt"),
    ("வோ", "Nth"),
    ("வொ", "nth"),
    ("வா", "th"),
    ("வி", "tp"),
    ("வீ", "tP"),
    ("வு", "T"),
    ("வூ", "T+"),
    ("வெ", "nt"),
    ("வே", "Nt"),
    ("வை", "it"),
    ("வ்", "t;"),
    ("வ", "t"),

    # ── ழ-வரிசை (Zha) ──
    ("ழௌ", "noo"),
    ("ழோ", "Noh"),
    ("ழொ", "noh"),
    ("ழா", "oh"),
    ("ழி", "op"),
    ("ழீ", "oP"),
    ("ழு", "O"),
    ("ழூ", "*"),
    ("ழெ", "no"),
    ("ழே", "No"),
    ("ழை", "io"),
    ("ழ்", "o;"),
    ("ழ", "o"),

    # ── ற-வரிசை (Ra - alveolar) ──
    ("றௌ", "nws"),
    ("றோ", "Nwh"),
    ("றொ", "nwh"),
    ("றா", "wh"),
    ("றி", "wp"),
    ("றீ", "wP"),
    ("று", "W"),
    ("றூ", "Wh"),
    ("றெ", "nw"),
    ("றே", "Nw"),
    ("றை", "iw"),
    ("ற்", "w;"),
    ("ற", "w"),

    # ── ஹ-வரிசை (Ha) ──
    ("ஹௌ", "n``"),
    ("ஹோ", "N`h"),
    ("ஹொ", "n`h"),
    ("ஹா", "`h"),
    ("ஹி", "`p"),
    ("ஹீ", "`P"),
    ("ஹெ", "n`"),
    ("ஹே", "N`"),
    ("ஹை", "i`"),
    ("ஹ்", "`;"),
    ("ஹ", "`"),

    # ── ஷ-வரிசை (Sha) ──
    ("ஷௌ", "n\\s"),
    ("ஷோ", "N\\h"),
    ("ஷொ", "n\\h"),
    ("ஷா", "\\h"),
    ("ஷி", "\\p"),
    ("ஷீ", "\\P"),
    ("ஷெ", "n\\"),
    ("ஷே", "N\\"),
    ("ஷை", "i\\"),
    ("ஷ்", "\\;"),
    ("ஷ", "\\"),

    # ── ஸ-வரிசை (Sa) ──
    ("ஸௌ", "n]s"),
    ("ஸோ", "N]h"),
    ("ஸொ", "n]h"),
    ("ஸா", "]h"),
    ("ஸி", "]p"),
    ("ஸீ", "]P"),
    ("ஸெ", "n]"),
    ("ஸே", "N]"),
    ("ஸை", "i]"),
    ("ஸ்", "];"),
    ("ஸ", "]"),

    # ── Uyir (vowels) ──
    ("அ", "m"),
    ("ஆ", "M"),
    ("இ", ","),
    ("ஈ", "<"),
    ("உ", "c"),
    ("ஊ", "C"),
    ("எ", "v"),
    ("ஏ", "V"),
    ("ஐ", "I"),
    ("ஒ", "x"),
    ("ஓ", "X"),
    ("ஔ", "xs"),

    # ── Special ──
    ("ஃ", "/"),

    # ── Standalone pulli (virama) ──
    ("்", ";"),
]

# Pre-compile a regex that matches any Tamil Unicode character (U+0B80–U+0BFF)
_TAMIL_UNICODE_RE = re.compile(r"[\u0B80-\u0BFF]")


def has_tamil_unicode(text: str) -> bool:
    """Return True if *text* contains any Tamil Unicode characters."""
    return bool(_TAMIL_UNICODE_RE.search(text))


def unicode_to_bamini(text: str) -> str:
    """Convert Unicode Tamil text to Bamini font encoding.

    Non-Tamil characters (English, digits, punctuation) pass through unchanged.
    The conversion replaces longest Unicode sequences first to correctly handle
    compound characters (uyir-mei letters).
    """
    if not text or not has_tamil_unicode(text):
        return text

    result = text
    for unicode_seq, bamini_seq in _UNICODE_TO_BAMINI:
        result = result.replace(unicode_seq, bamini_seq)
    return result


# Reverse mapping for Bamini → Unicode.
# Sort by length of the Bamini sequence in descending order to avoid partial matches first.
_BAMINI_TO_UNICODE: list[tuple[str, str]] = sorted(
    [(bamini_seq, unicode_seq) for unicode_seq, bamini_seq in _UNICODE_TO_BAMINI],
    key=lambda x: len(x[0]),
    reverse=True
)


def bamini_to_unicode(text: str) -> str:
    """Convert Bamini legacy font encoded text to Unicode Tamil."""
    if not text:
        return text

    result = text
    for bamini_seq, unicode_seq in _BAMINI_TO_UNICODE:
        result = result.replace(bamini_seq, unicode_seq)
    return result


def convert_values_to_bamini(values: dict[str, str]) -> dict[str, str]:
    """Convert all Tamil Unicode values in an extraction row dict to Bamini."""
    return {key: unicode_to_bamini(val) for key, val in values.items()}


def convert_rows_to_bamini(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    """Convert all Tamil Unicode values in a list of extraction rows to Bamini."""
    return [convert_values_to_bamini(row) for row in rows]
