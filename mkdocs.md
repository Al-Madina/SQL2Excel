site_name: SQL2Excel
site_description: Exporting SQL queries and DataFrames to Excel with charts and data visualization.
site_author: Ahmed Hassan
repo_url: https://github.com/Al-Madina/SQL2Excel
repo_name: Al-Madina/SQL2Excel

theme:
  name: material
  palette:
    # Light mode
    - media: "(prefers-color-scheme: light)"
      scheme: default
      primary: indigo
      accent: indigo
      toggle:
        icon: material/brightness-7
        name: Dark mode
    # Dark mode
    - media: "(prefers-color-scheme: dark)"
      scheme: slate
      primary: indigo
      accent: indigo
      toggle:
        icon: material/brightness-4
        name: Switch to light mode
  features:
    - navigation.instant
    - navigation.tracking
    - navigation.tabs
    - navigation.tabs.sticky
    - navigation.sections
    - navigation.top
    - search.suggest
    - search.highlight
    - search.share
    - content.code.copy
    - content.code.annotate

plugins:
  - search
  - mkdocstrings:
      handlers:
        python:
          options:
            docstring_style: numpy  
            show_root_heading: true
            show_symbol_type_heading: true
            show_symbol_type_toc: true
            merge_init_into_class: true

markdown_extensions:
  - admonition
  - pymdownx.details
  - pymdownx.superfences:
      custom_fences:
        - name: mermaid
          class: mermaid
          format: !!python/name:pymdownx.superfences.def_mermaid
  - pymdownx.highlight:
      anchor_linenums: true
      line_spans: __span
      pygments_lang_class: true
  - pymdownx.inlinehilite
  - pymdownx.snippets
  - pymdownx.tabbed:
      alternate_style: true
  - pymdownx.emoji:
      emoji_index: !!python/name:material.extensions.emoji.twemoji
      emoji_generator: !!python/name:material.extensions.emoji.to_svg
  - attr_list
  - md_in_html
  - tables
  - toc:
      permalink: true

nav:
  - Home: index.md
  - Getting Started:
      - Installation: installation.md
      - Quickstart: quickstart.md
  - User Guide:
      - SQL to Excel: guide/sql-export.md
      - DataFrames & Charts: guide/charts.md
      - Database Connectors: guide/databases.md
      - CLI Reference: guide/cli.md
  - API Reference:
      - Core API: api/core.md
      - CLI API: api/cli.md
  - Project Info:
      - Contributing: contributing.md
      - Changelog: changelog.md
      - License: license.md