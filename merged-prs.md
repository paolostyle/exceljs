# Included fixes
ExcelJS has a ton of users and even though it's essentially abandoned I want to include the contributed fixes that make sense. I try to merge commits from original contributors, but if the PR quality is low (this includes e.g. bad formatting) I implement the fixes by myself instead.

### Legend
- ✅ - merged
- 🔧 - manually applied
- 🚧 - need to review
- ❓ - unsure
- ❌ - rejected

### PRs open, state at 2025-02-01

- ✅ [Make to work with expressions with no formulae (#2883)](https://github.com/exceljs/exceljs/pull/2883) 
- ✅ [Fix error of adding image in certain situations (#2876)](https://github.com/exceljs/exceljs/pull/2876)
- 🚧 [Fix: Changed error-prone race conditions (#2874)](https://github.com/exceljs/exceljs/pull/2874)
- ❌ [Bumping unzipper (DUPLICATED #2744) (#2869)](https://github.com/exceljs/exceljs/pull/2869)
- 🚧 [Introducing styleCacheMode. Up to 3x performance improvements on xlsx… (#2867)](https://github.com/exceljs/exceljs/pull/2867)
- ❌ [fix: add check for empty target on worksheet-xform reconcile (#2852)](https://github.com/exceljs/exceljs/pull/2852)
- ❓ [fix boolean read val error like as ：\<strike val="0"/> (#2851)](https://github.com/exceljs/exceljs/pull/2851)
- ❓ [feat: support web-native streams for read/write methods (#2849)](https://github.com/exceljs/exceljs/pull/2849)
- ❌ [Shuntagami patch 1 (#2847)](https://github.com/exceljs/exceljs/pull/2847)
- ❓ [Update xlsx.js to allow compat with non office generated files. (#2846)](https://github.com/exceljs/exceljs/pull/2846)
- ❌ [update-dependency-version (#2812)](https://github.com/exceljs/exceljs/pull/2812)
- 🚧 [Added quote prefix feature (#2809)](https://github.com/exceljs/exceljs/pull/2809)
- [place pageSetUpPr in the end of sheetPr to fix getting broken xlsx do… (#2807)](https://github.com/exceljs/exceljs/pull/2807)
- [Fix corrupted file with conditional formatting and hyperlinks (#2803)](https://github.com/exceljs/exceljs/pull/2803)
- [fix: worksheet-reader hidden prop (#2800)](https://github.com/exceljs/exceljs/pull/2800)
- [Issue 2790/xlsx stream missing worksheets (#2791)](https://github.com/exceljs/exceljs/pull/2791)
- [:memo: Fix errors in document about image embedding (#2783)](https://github.com/exceljs/exceljs/pull/2783)
- [Add option to set height/width of notes (#2782)](https://github.com/exceljs/exceljs/pull/2782)
- [fix: setting cell style attribute clones style object (#2781)](https://github.com/exceljs/exceljs/pull/2781)
- [improve: add logs to help developers troubleshoot issues (#2779)](https://github.com/exceljs/exceljs/pull/2779)
- [Bug2675 table creation accepts invalid names (#2767)](https://github.com/exceljs/exceljs/pull/2767)
- [fix #2751 - Csv reading - cells filled with spaces only are converted to 0 (#2752)](https://github.com/exceljs/exceljs/pull/2752)
- [bump: Bumping unzipper to mitigate license issue (#2744)](https://github.com/exceljs/exceljs/pull/2744)
- [Don't render empty rich text substrings (#2737)](https://github.com/exceljs/exceljs/pull/2737)
- [Improve conditional formatting settings (#2736)](https://github.com/exceljs/exceljs/pull/2736)
- [[Doc]: Improve readme.md (#2733)](https://github.com/exceljs/exceljs/pull/2733)
- [Fix type mismatch in Address interface (#2720)](https://github.com/exceljs/exceljs/pull/2720)
- [fix: add proper version control to deps (#2710)](https://github.com/exceljs/exceljs/pull/2710)
- [Fix date parsing for Strict OpenXML spreadsheets (#2702)](https://github.com/exceljs/exceljs/pull/2702)
- [Issue: style.xml has [Object object] as formatCode (#2698)](https://github.com/exceljs/exceljs/pull/2698)

This will be updated further