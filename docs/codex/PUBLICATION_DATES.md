# Publication dates and ordering (UPD-0204)

## Selection

RSS items use their direct, unnamespaced `pubDate`. Atom entries use their direct Atom-namespace `published`, with `updated` only when published is absent or blank. A present malformed publication date stays undated; an edit timestamp cannot silently replace it. Foreign elements with matching date names do not control selection. The established publisher-ID extraction remains unchanged to preserve existing episode/history identities, including its legacy local-name compatibility.

`Get-EpisodeData` returns `PubDate` as a UTC `DateTimeOffset` or null. `PubDateOriginal` and `PubDateSource` retain the selected original text and field in memory. They are not serialized into history or diagnostic exports. Invalid/missing dates do not discard titles, enclosure URLs or publisher IDs.

`Get-OrderedPodcastEpisode` compares UTC ticks numerically, newest first. Dated entries precede undated entries, including the minimum supported date. Equal instants and all undated entries retain their supplied source order using explicit numeric positions; this does not rely on engine sort stability or culture collation. A feed that rearranges tied or undated entries may therefore change their selection order. Latest takes one, Custom takes the requested count, and All retains the full ordered snapshot. Collection outputs remain arrays and source objects are not modified.

## Accepted date text

The bundled parser uses numeric Gregorian components and explicit offsets, never a machine-culture parse or current-date default. Maximum input length is 256 characters, with 250 ms per regex match. Leading/trailing whitespace is trimmed.

| Form | Policy |
| --- | --- |
| `yyyy-MM-ddTHH:mm:ssZ` or numeric offset | Seconds required; `+HH:mm` and `+HHmm` accepted. Lowercase t/z and a space separator are compatibility forms. |
| Fractional seconds | One through seven digits, preserved at .NET tick precision. More precision is rejected rather than silently rounded. |
| Date-only or ISO timestamp without zone | Compatibility forms interpreted as UTC, midnight for date-only. They do not depend on the computer's time zone. |
| English RSS/RFC-style date | Optional three-letter weekday plus comma; one/two-digit day, English three-letter month, two/four-digit year, hour/minute and optional seconds, explicit zone required. The weekday is a compatibility hint and is not checked against the numeric date. |
| Two-digit RSS year | Fixed mapping: 00-49 means 2000-2049; 50-99 means 1950-1999. No calendar TwoDigitYearMax dependence. |
| Named RSS zone | UT/UTC/GMT/Z mean UTC; EST/EDT/CST/CDT/MST/MDT/PST/PDT use their fixed historical offsets, not live daylight-saving rules. |

Numeric offsets must have minutes 00-59 and magnitude no greater than 14:00. Negative zero is treated as UTC for chronological comparison. Impossible calendar dates, out-of-range UTC instants, leap seconds, time-only input, localized/ambiguous slash dates, arbitrary zone names, military zones and RFC comments/folding are unsupported and return null. This is a documented feed-date subset with compatibility extensions, not full validation of every RSS/Atom date production. Atom's standard date construct requires a full RFC3339 timestamp; date-only, zone-less and lowercase forms above are explicit project choices [S9, S22-S23 in SOURCES.md].

## New filenames and established archives

New filenames use the **UTC calendar date**, `yyyy-MM-dd`, followed by the existing title and full identity suffix. For example `2026-09-01T00:30:00+02:00` names a new file with `2026-08-31`. Undated entries omit the date. Existing path-length budgeting may also omit the date. Metadata comparison serializes the UTC instant consistently so equivalent offsets do not produce a false duplicate-identity conflict.

History remains authoritative: an existing stable episode identity keeps its recorded relative path across date parsing, date metadata, title or signed-URL changes. No old file is renamed or overwritten because a date improves. History schemas, feed/episode fingerprints, completion evidence and resume ownership are unchanged.

`LegacyPubDate` is a separate lookup-only compatibility hint with the earlier `pubDate -> updated -> published` priority. Explicitly zoned values use the current computer's local calendar date, reproducing the old naming convention on that computer; date-only/zone-less ISO values keep their original wall date. The historical filename helper accepts older in-memory DateTime inputs and empty dates. This hint never controls chronological selection, new names or recorded destinations, and never establishes verified/adopted evidence. An old archive created in another time zone, with unsupported date text, or with different prior metadata can still require explicit manual review. Unknown files and originals remain preserved.

## Validation boundaries

Unit regressions exercise en-US, de-DE and fi-FI, offsets, precision/range limits, missing/tied order, Atom precedence, provenance, retained IDs and historical hints. Loopback tests exercise actual Latest/Custom downloads, UTC midnight filenames, write-free preview and stable recorded destinations/bytes after date changes. No real archive or publisher feed is used. Media choice/format, pagination, launchers and release behavior remain later tasks.
