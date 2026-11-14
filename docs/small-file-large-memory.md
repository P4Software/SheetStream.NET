---
title: "A 98 KB spreadsheet that takes 400 MB"
description: "How a 98 KB .xlsx file could drive a reader to 400 MB, what the 0.5.0 fix missed, and what changed in 0.6.0 to stop it."
---

# A 98 KB spreadsheet that takes 400 MB, and what it took to stop it

*November 2026, written after the 0.6.0 release.*

An `.xlsx` file is a zip archive full of XML. Zip compresses repetition very well: a hundred million copies of the letter A fit in under 100 KB. So a spreadsheet that arrives as a small upload can ask the server that opens it for gigabytes.

SheetStream.NET writes and reads Excel files by streaming. In version 0.5.0 we fixed a report of exactly this: a 192 KB workbook with one cell of 200 million characters drove a reading process to 1.5 GB. We added limits, wrote tests, published an advisory, and believed the problem was closed.

It was not. This is what we got wrong, how we found out, and what we do differently now. The advisory is [GHSA-rvj5-c489-67jr](https://github.com/P4Software/SheetStream.NET/security/advisories/GHSA-rvj5-c489-67jr); everything described here is fixed in 0.6.0.

## What we got wrong

We had fixed the attack we were shown, not the class it belonged to. The fix counted the bytes between one `<` and the next, on the reasoning that a tag or a text value cannot contain a raw `<`, so a long gap means a long value. That is true, and it left at least seven other ways in:

1. **CDATA.** Inside `<![CDATA[ ... ]]>` a `<` is ordinary text. A section made of a hundred million `<` characters never produced a long gap, and the XML parser buffers a CDATA section whole. A 98 KB file took 418 MB.
2. **UTF-16.** We counted bytes. In UTF-16 the character U+3C3C is two bytes that both look like `<`. A long attribute made of that character was invisible to the check.
3. **Cells without addresses.** A cell normally says where it is (`r="B7"`), and we refused columns past Excel's last one. A cell may also omit its address and mean "the next column". We did not count those. A 524 KB file made the reader hold 313 MB for one row.
4. **Names.** An XML parser remembers every distinct element and attribute name for as long as it reads a document. A sheet made of `<a1/><a2/><a3/>...` grows that table without limit, however small each tag is.
5. **Many small attributes.** Our limit on a tag was sized for text. A tag costs the parser far more than text of the same length, because it keeps a record per attribute: a tag with a very large number of them cost hundreds of megabytes while staying under the limit.
6. **The other ways of opening a file.** The library can also open a workbook to edit it or to fill it as a template. We had documented those as "for files you own" and given them no limits at all. A 429 KB file reached 5.6 GB, and a 16 KB file of nested elements was still running after five minutes when we stopped it.
7. **What a parsed part costs.** We capped the size of the parts we keep in memory. A cell format is eighteen bytes of XML and well over a hundred as an object. The cap on the file said little about the memory.

None of these needed a clever attacker. Each is one line of a script that writes XML.

## How we found out

By writing down the attacks as a table instead of as tests.

One axis is every hostile shape we could think of: about sixty of them now. The other axis is every way the library opens a workbook: the reader, the async reader, the data-table reader, merged ranges, hyperlinks, editing, template filling, template copying. Every shape runs through every entry point. Adding a shape or an entry point adds a whole row or column, so nobody has to remember which pairs matter.

Three checks run on that table, because each catches something the others cannot:

- **The outcome of every pair is recorded** in a text file in the repository, as "reads" or "refused". A change that turns one into the other fails the build until someone looks at it. "It did not crash" is too weak a test.
- **A refusal must be cheap.** If refusing a file allocates as much as the file would have, the limit is decoration.
- **The whole table runs again in a separate process whose heap the runtime caps at 64 MB.** This is the one that matters. Allocation counters cannot tell a 40 MB string from 40 MB of short-lived rows. A hard cap can: anything that holds the hostile content, even for a moment, dies with an out-of-memory error that names the pair. The table passes at 40 MB.

To check that the checks themselves work, we disabled each guard in turn and confirmed the capped process failed and named the right pair.

## Then we had it attacked separately

The table was written alongside the fix, which is the weakness of every test suite: it knows what the fix expects. So before release the work went through a separate adversarial review pass that started from the code, not from our tests, with one instruction: break it, and prove it with a file.

The first pass came back with thirteen confirmed problems, and a second pass found more. The pattern was instructive: we had bounded memory and forgotten time.

- A relationship target made of `./` repeated, six kilobytes zipped, took 43 seconds to open. We normalised paths by removing segments from the front of a list.
- A zip directory can point many names at one compressed block. Saving an edited copy unpacked and repacked every one: a 259 KB file cost 10 GB of work, and the format allows far more.
- A very long number format was inspected again for every cell that used it: a 10 KB file took 18 seconds.
- `ReadDataTable` builds a table of rows times columns. Three thousand rows with one cell each in the last column, 15 KB on disk, became 1 GB.
- A formula from a file, nested a few thousand levels deep, overflowed the stack of the formula evaluator. A stack overflow cannot be caught. It ends the process.

Every one is fixed and has a regression test built from the file that broke it. The review's own exploit harness and fuzzer were then rerun against the fixes.

## What we do differently

- **Decode before you check.** The guard now runs on characters, after decoding, so no encoding can disguise what the parser will see.
- **Charge what you keep.** Everything that outlives the read of one row is charged, in bytes, to one budget per workbook. The limit is now a statement about memory, not about files.
- **Bound time as well as memory.** Streamed parts have a size limit too, and anything that costs more than the size of what it reads is treated as a bug.
- **Compare with Excel.** We built a corpus of workbooks written by Excel, with Excel's own account of every cell, and compare our reader with it. It found a bug on its first run: Excel writes a second element named `workbookPr` in an extension list, and ours let it overwrite the first, so every genuine Excel file in the 1904 date system had been read four years off. Our own files never had that element, which is why our own tests passed.
- **Fuzz from real files.** Mutating a workbook Excel wrote reaches code that mutating our own output cannot.
- **Have it attacked before release, in a pass that does not start from the fix's own tests.**

## The honest summary

With default settings, files built to exhaust memory that used to take up to 5.9 GB now peak at 65 MB or less, through every way of opening them. What remains is by design and documented in [Reading workbooks](reading-workbooks.md): a workbook may make the library hold up to its memory limit (256 MB by default, which an upload endpoint should lower), a streamed sheet may be as large as `MaxStreamedPartBytes` allows (4 GB by default, which an upload endpoint should also lower), and the packed file itself is your web framework's to cap.

We shipped an incomplete fix once. The thing we would tell another maintainer is not "add limits". It is: write the attacks down as a table, run the table under a memory cap, and have it attacked separately before your users do it for you.
