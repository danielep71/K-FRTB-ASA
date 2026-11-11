# Wiki evidence update

The [prepared patch](2026-11-01-excel-evidence.patch), part of #32, corrects
eight wiki pages that still say no desktop Excel run exists. It is based on
wiki commit `d53a0e034b358eb7031bdd3f16882bfd6174c592` and the published
[October 31 Excel record](https://github.com/danielep71/K-FRTB-SA/blob/4a3b97412e7a3b7f4c80b736d239a88377f851d5/docs/evidence/2026-10-31-excel-run-323260b.md).

The run covers candidate `323260bb38116cd64c616e7acd73af7d8bdc9482` on
`release/v0.3.0`, on one Windows/Excel host. The patch preserves the remaining
manual checks, the private-bundle limitation and the distinction between
Excel, LibreOffice and regulatory validation. It does not change the dated
October 27 evaluation. The layout fix #203 is closed, but no Excel run has
seen it yet, so the pages leave that check to the all-sheet review #206.

The project wiki is stored in its own Git repository, `K-FRTB-SA.wiki`. To
publish the update, apply the patch to a clean wiki checkout at the named
commit with `git apply --check` followed by `git apply`, inspect the diff and
push it. If the wiki has advanced, reconcile the patch against its new head
first. #32 records whether the update is published.

Validation: the patch applies to the named wiki commit, the resulting diff
has no whitespace errors, and the eight pages retain their local page and
anchor targets. No Excel run or private-bundle validation was repeated.
