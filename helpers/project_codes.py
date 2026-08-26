"""Registry ID -> short project code for the Master Inventory ``Inventory ID``.

MANUALLY-SYNCED COPY. The authoritative version lives at
``ACCU_reporting/helpers/project_codes.py`` -- change it there first, then propagate here,
the same convention ``inv_datadict.py`` already follows. These repos share no installable
package, so a divergent copy mints inconsistent internal references with no error.


``Inventory ID`` is the human-readable internal reference carried on every MI row:
``<f|a>-<code>-<yymmdd>-<amount>`` (``f-`` = forecast reference, ``a-`` = actual). Every generator
derived ``<code>`` as ``project_name[:4].title()``, which is not unique across the portfolio -- and
where two projects share a prefix the internal reference stops identifying the project at all.

**This is not hypothetical.** Reconciling ``submission_approval_raw.csv`` against the Carbon
Delivery Hub in August 2026 turned up three issuance rows keyed to the wrong project, and the
shared ``Paro`` prefix is precisely why the error survived review: the exception report's own
caution read *"both projects use the same ``Paro`` inventory prefix, so the Inventory ID does not
discriminate between them -- the CER application record is the only arbiter."* An internal reference
that cannot tell two projects apart forces every attribution question out to CER paperwork.

Decided 24 August 2026 (Georgina): **Paroo North and Paroo South take distinct codes, PRN and PRS.**

Known prefix collisions in the portfolio as at 24 August 2026::

    Paro  ERF104646  Paroo River North Environmental Project      -> PRN
    Paro  ERF104559  Paroo River South Environmental Project      -> PRS
    West  EOP101162  Western Farm Trees                           -> no code assigned yet
    West  ERF101667  Westmere Regeneration Project                -> no code assigned yet
    West  ERF112715  Westerton Regeneration Project               -> no code assigned yet

The three ``West`` projects are the same defect and still need codes; they are deliberately left
unassigned rather than invented here, because a code baked into an internal reference is data, not
formatting, and CCG should choose it. Until they are assigned they continue to fall back to
``West``, exactly as today -- adding a project to :data:`PROJECT_CODES` is the whole fix.

**Codes apply to newly minted references only.** Existing ``Paro`` references in the Master
Inventory are not rewritten: the MI merge ``Key`` is a hash over Name + Registry ID + RP dates +
Status + Total Amount and does *not* include the Inventory ID, so a code change cannot break the
merge -- but the Inventory ID is used for cross-referencing rows to forecasts by hand, and silently
rewriting historical references would break that. New rows get the new code; old rows keep theirs.

This module is the authoritative copy. ``HIR_Forecasting/scripts/project_codes.py`` and
``EP_Forecasting/helpers/project_codes.py`` are manually-synced duplicates, following the same
convention as ``inv_datadict.py`` -- these repos share no installable package. **Change this file
first, then propagate.**
"""

from __future__ import annotations

import re

#: Registry ID -> internal reference code. Registry ID is the key because it is the only project
#: identifier that is stable and unique; ``Project Name`` is neither, which is what this fixes.
PROJECT_CODES: dict[str, str] = {
    "ERF104646": "PRN",  # Paroo River North Environmental Project
    "ERF104559": "PRS",  # Paroo River South Environmental Project
}

#: Code used when neither an explicit mapping nor a usable project name is available.
DEFAULT_CODE: str = "Proj"

#: Length of the legacy name-derived fallback code. Explicit codes are not held to it -- ``PRN``
#: is 3 characters and correct; the format has never been fixed-width and nothing parses it
#: positionally (the delimiter is ``-``).
FALLBACK_LENGTH: int = 4

#: Legacy name-prefix overrides that predate this module, preserved so existing references keep
#: minting unchanged. ``EP_Forecasting`` applied these inline; they are keyed on a lower-cased
#: ``startswith`` match against the project name.
LEGACY_NAME_PREFIXES: dict[str, str] = {
    "big creek": "Big ",
    "cpc beef herd": "CPC ",
}


def project_code(registry_id: str | None, project_name: str | None) -> str:
    """Return the internal reference code for a project.

    Resolution order -- Registry ID first, because it is the only unique identifier:

    1. an explicit :data:`PROJECT_CODES` entry for ``registry_id``;
    2. a :data:`LEGACY_NAME_PREFIXES` match on ``project_name``, so references already in the
       Master Inventory keep minting the same way;
    3. the legacy fallback, ``project_name``'s first :data:`FALLBACK_LENGTH` letters, title-cased;
    4. :data:`DEFAULT_CODE` when the name yields nothing.

    Args:
        registry_id: CER registry ID, e.g. ``"ERF104559"``. May be blank or ``None``.
        project_name: Master Inventory ``Project Name``. May be blank or ``None``.

    Returns:
        The code to embed in an ``Inventory ID``.

    Examples:
        >>> project_code("ERF104559", "Paroo River South Environmental Project")
        'PRS'
        >>> project_code("ERF104646", "Paroo River North Environmental Project")
        'PRN'
        >>> project_code("ERF101667", "Westmere Regeneration Project")
        'West'
        >>> project_code(None, None)
        'Proj'
    """
    key = str(registry_id or "").strip()
    if key in PROJECT_CODES:
        return PROJECT_CODES[key]

    name = str(project_name or "").strip()
    lowered = name.lower()
    for prefix, code in LEGACY_NAME_PREFIXES.items():
        if lowered.startswith(prefix):
            return code

    letters = "".join(re.findall(r"[A-Za-z]", name))[:FALLBACK_LENGTH].title()
    return letters or DEFAULT_CODE
