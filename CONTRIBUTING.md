# Contributing to openEfficiency

Bug reports, documentation corrections and focused pull requests are welcome. openEfficiency is maintained by an individual; review and support depend on availability.

## Before opening an issue

Search [existing issues](https://github.com/FabPei/oPenEfficiency/issues) and consult the [wiki](https://github.com/FabPei/oPenEfficiency/wiki). Include the add-in version, Office version and bitness, Windows version, steps to reproduce, expected result and actual result. A small, non-confidential sample is useful when the problem depends on a presentation.

Do not post credentials, confidential slides or unredacted logs. See the [Code of Conduct](CODE_OF_CONDUCT.md) for community expectations.

## Suggesting a feature

Explain the task you are trying to complete, the current workaround and the expected result. A feature request is not a commitment to implement it or a release date.

## Pull requests

1. Start from the public branch or release tag relevant to the change. Keep the patch focused and explain its purpose.
2. Follow the nearby C#/WPF conventions. Use the existing PowerPoint access wrappers and take care with selection, COM object lifetimes, startup and shutdown.
3. Keep command names, descriptions, tooltips and relevant documentation consistent. Update user-facing documentation when behavior changes.
4. Describe the checks you actually performed and any checks you did not perform. For documentation-only changes, check links and examples; do not claim an Office test or build that was not run.
5. Follow the [Code of Conduct](CODE_OF_CONDUCT.md).

See [Development](https://github.com/FabPei/oPenEfficiency/wiki/Development) for the public source layout and local prerequisites. A local development build and a packaged release are different artifacts.

## Licensing contributions

The project uses **GNU AGPL v3**; read [LICENSE](LICENSE) and the [licensing guide](docs/licensing.md). Submit only material you have the right to contribute under the applicable terms. Preserve third-party authorship, copyright and license notices; document any new external source and its license.

This contribution guide does not transfer copyright ownership or revoke independent licenses already validly granted. Contributions and software are supplied subject to their applicable license terms, without a promise of commercial support.
