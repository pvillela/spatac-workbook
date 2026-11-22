# Installing the tools

Two command-line tools come with the workbook:

- [`xlsx-diff`](xlsx-diff.md) compares two versions of a workbook, any workbook.
- [`spatac-investment-plan`](investment-plan.md) extracts the Investment Plan from the
  Spatac workbook.

They run in a terminal: Command Prompt or PowerShell on Windows, a shell on Linux.

## Download

Each release on the project's
[GitHub releases page](https://github.com/pvillela/spatac-workbook/releases) has two downloads:

- **Windows:** `spatac-workbook-x86_64-pc-windows-msvc-….zip`. Unzip it, for example into
  `C:\Tools\spatac`.
- **Linux:** `spatac-workbook-x86_64-unknown-linux-gnu-….tar.gz`. Extract it with
  `tar -xzf …`. It runs on Ubuntu 26.04 and later.

Each download holds both tools, the licences and the third-party notices. The file names carry
the version and the commit they were built from.

To run a tool from any folder, add its folder to your `PATH`, or rename the files to
`xlsx-diff` and `spatac-investment-plan` (`.exe` on Windows) and put them in a folder already
on it. Otherwise, give the full path each time.

## Building from source

With Rust installed:

```
cargo build --release
```

The tools are written to `target/release/`.

## Help

Every tool explains itself:

```
xlsx-diff --help
spatac-investment-plan --help
```
