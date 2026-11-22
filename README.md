# spatac-workbook

Utilities and documentation to support Antiquary reserve fund project tracking and investment
planning.

The Spatac workbook is an Excel workbook, also used in LibreOffice Calc 26.2. With it, the board
and SPATAC (Special Projects and Technical Advisory Committee) track reserve fund projects and plan
investments. This repository holds:

- **`xlsx-diff`**, a command-line tool that compares two versions of a workbook (any workbook, not
  only Spatac's) and writes the differences as Markdown.
- **`spatac-investment-plan`**, a command-line tool that extracts the Medium_Term_Cashflow and
  Schedule_of_Investments tabs into a separate InvestmentPlan workbook.
- The documentation for using the workbook and the tools.

The workbooks themselves are not in this repository.

## Documentation

The guide to the workbook and the tools is published at
<https://pvillela.github.io/spatac-workbook/>. Its sources are the Markdown files in `docs/`,
built with [mdBook](https://rust-lang.github.io/mdBook/):

```
mdbook build        # writes book/
mdbook serve        # previews at http://localhost:3000
python3 scripts/check-doc-links.py
```

`docs/archive/` holds the design record, for maintainers. It is not part of the published book.

## Building

```
cargo build --release
```

The binaries are written to `target/release/`. Release builds need generated third-party notices
first:

```
bash scripts/gen-notices.sh
```

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE))
- MIT license ([LICENSE-MIT](LICENSE-MIT))

at your option.
