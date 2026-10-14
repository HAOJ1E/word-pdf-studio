# WordPdfStudio verification plan

## Input fixtures

Run `build_fixtures.py` with the bundled Python runtime. `fixtures.json` lists all expected text markers and page counts. All fixture content is synthetic and all fixtures remain internal QA assets.

- `中文 空格路径/双语报告 2026.docx`: expected two portrait A4 pages, Chinese and English, numeric symbols, table, inline chart, header, footer, page-number field, explicit page break, bullets and styled runs.
- `duplicate names/客户甲/相同名称.docx` and `duplicate names/客户乙/相同名称.docx`: distinct text markers; batch conversion must keep both outputs and prevent accidental overwrite.
- `landscape.docx`: one landscape page with image; preserve page dimensions.
- `empty valid.docx`: valid blank input, expected one blank page.
- `中文 空格路径/富文本 sample.rtf`: two pages with Unicode escapes and explicit page break.
- Broken ZIP, missing Word body and unsupported extension inputs must fail with useful messages while later queued jobs still run.
- Three PNG/JPEG inputs cover landscape, portrait and transparency.

## Conversion checks

1. Invoke the actual packaged executable with documented command-line options, using absolute input/output paths containing spaces and Chinese. Verify PDF signature, expected pages and text, image presence, header/footer and page-specific markers using `verify_outputs.py`.
2. Render the main output at 110 dpi using bundled Poppler. Inspect every page for missing Chinese glyphs, clipping, distorted images, wrong table widths, header/footer overlap and unexpected page changes.
3. Repeat a conversion to an existing destination and verify collision policy or explicit overwrite handling; never accept silent data loss.
4. Verify temporary engine profiles are isolated, that concurrent or consecutive jobs succeed, and that cancellation closes only the application's own converter process. Check filenames and logs avoid passwords.

## PDF utility checks

1. Merge main (2 pages) and duplicate-alpha (1 page); assert three pages and exact marker order.
2. Extract pages `2,1` or supported ordering and assert documented order; validate invalid, duplicate, zero, negative and out-of-bounds page selections produce actionable messages.
3. Split a two-page PDF into individual files; assert page count one and matching marker in each.
4. Rotate 90/180/270 degrees; inspect `/Rotate` and render to verify visible orientation. If implementation rotates content instead of `/Rotate`, use rendered orientation as the final authority.
5. Apply Latin and Chinese watermark text; verify text and rendered placement, opacity and visible Chinese characters on all pages.
6. Encrypt with distinct user/owner passwords; assert encryption flag, wrong-password rejection, correct password opening and preservation of text/page count. Confirm GUI does not retain passwords in history or logs.
7. Image-to-PDF: image ordering, one page per image, portrait/landscape aspect preservation and transparent-image background. Render and inspect all pages.
8. Reject malformed/encrypted-without-password PDFs and verify task isolation.

## GUI acceptance

Check Windows scaling at 100%, 125% and 150%; narrow window/minimum size; drag/drop and file picker; keyboard focus; queue progress; cancel; retry; output-folder opening; output preview; clear success/failure counts; local settings persistence. Conversion and utility work must leave the UI responsive. Avoid relying on color alone for status.

## Portable and clean-machine validation

The strongest evidence is a fresh Windows x64 VM with no Microsoft Office, LibreOffice, .NET runtime or Python installed and network disconnected. Copy only the release artifact, unpack into a user-writable Chinese/space path and run as a standard user. Verify bundled converter path, font availability, conversion and every PDF utility; reboot and repeat. Also test a read-only installation directory with writable user temp/output folders. Check the artifact's dependency manifest and license notices.

Local conversion success proves application functionality on this host; it does not independently prove a clean computer has all native DLLs, fonts, the Visual C++ runtime or the correct architecture. Report clean-machine testing as pending unless performed on an actual isolated environment. A self-contained .NET publish does not itself bundle the LibreOffice engine or its native dependencies.

## Anticipated limits

Word and LibreOffice may paginate complex documents differently. Fonts, embedded fonts, equations, floating objects, tracked changes, macros and legacy `.doc` formats need broader corpus testing. PDF/A conversion must be validated by a dedicated PDF/A checker if advertised. Encryption implementation must identify supported algorithms and limits. Huge files and pathological documents require timeout/cancellation and disk-space checks. A portable folder cannot be replaced by shipping a small launcher EXE alone; all bundled files must be distributed together unless the build intentionally packages the complete engine.
