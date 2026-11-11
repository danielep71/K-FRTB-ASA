# Wiki updates

The [prepared patch](2026-11-01-excel-evidence.patch), part of #32, corrects
eight wiki pages that still say no desktop Excel run exists. It is based on
wiki commit `d53a0e034b358eb7031bdd3f16882bfd6174c592` and the published
[October 31 Excel record](https://github.com/danielep71/K-FRTB-ASA/blob/4a3b97412e7a3b7f4c80b736d239a88377f851d5/docs/evidence/2026-10-31-excel-run-323260b.md).

The run covers candidate `323260bb38116cd64c616e7acd73af7d8bdc9482` on
`release/v0.3.0`, on one Windows/Excel host. The patch preserves the remaining
manual checks, the private-bundle limitation and the distinction between
Excel, LibreOffice and regulatory validation. It does not change the dated
October 27 evaluation. The layout fix #203 is closed, but no Excel run has
seen it yet, so the pages leave that check to the all-sheet review #206.

The [rename patch](2026-11-11-rename.patch), part of #261, replaces the
former name K-FRTB-SA with K-FRTB-ASA on the 19 wiki pages that use it,
including the links to the repository and to the website, whose address
GitHub does not redirect. It applies on top of the evidence patch.

The project wiki is stored in its own Git repository, `K-FRTB-ASA.wiki`. To
publish both updates, apply the evidence patch and then the rename patch to a
clean wiki checkout at the named commit, each with `git apply --check`
followed by `git apply`, inspect the diff and push it. If the wiki has
advanced, reconcile the patches against its new head first. #32 and #261
record whether the updates are published.

Validation: the evidence patch applies to the named wiki commit and the
rename patch applies on top of it; the resulting diffs have no whitespace
errors, and the pages retain their local page and anchor targets. No Excel run or private-bundle validation was repeated.
