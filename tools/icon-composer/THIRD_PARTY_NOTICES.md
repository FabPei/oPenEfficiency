# Third-party sources in the authoring tool

Maintainer-authored editor, Composer and exporter code follows the existing root
[AGPL v3 license](../../LICENSE). No project license is changed by this delivery.
The following separately licensed inputs keep their original terms and notices.

| Input | Terms and retained provenance |
| --- | --- |
| Lucide and inherited Feather sources | ISC and the inherited MIT notices in [LUCIDE-LICENSE.txt](kit/notices/LUCIDE-LICENSE.txt); [upstream license](https://github.com/lucide-icons/lucide/blob/main/LICENSE). |
| Microsoft Fluent sources | [Retained MIT notice](kit/notices/FLUENT-LICENSE.txt); [upstream license](https://github.com/microsoft/fluentui-system-icons/blob/main/LICENSE). |
| Material Design Icons | [Pictogrammers notice](kit/notices/MDI-LICENSE.txt), [Apache 2.0 text](kit/notices/APACHE-2.0.txt), and pinned URLs/hashes in kit provenance; [upstream terms](https://pictogrammers.com/docs/general/license/). |
| Noto Sans number outline graphic | [OFL 1.1 and copyright notice](kit/notices/NOTO-OFL.txt), [pinned outline provenance](kit/provenance/noto-number-outlines.json); [official OFL text](https://openfontlicense.org/open-font-license-official-text/). No font binary is included. |
| jsdom and tsx | MIT, installed from the exact npm lockfile; retain their package notices if distributing their code. |
| @resvg/resvg-js | MPL-2.0, installed unchanged from the exact npm lockfile; [MPL text](https://www.mozilla.org/en-US/MPL/2.0/). Its source and original terms remain available through the package repository. |
| TypeScript and type declarations | Development-only pinned npm dependencies; their package licenses are recorded in the lockfile. |

No node_modules, third-party runtime source, font files or graphical-editor libraries
are copied into the public source bundle. npm installs its dependencies separately.
The public kit omits the Font and Language entries: the existing evidence does not
establish a grant to redistribute their individual Windows-font glyph outlines as
reusable building blocks. This is a bounded source-delivery decision, not a conclusion
about other use of those fonts. See the [Microsoft font FAQ](https://learn.microsoft.com/en-us/typography/fonts/font-faq).
