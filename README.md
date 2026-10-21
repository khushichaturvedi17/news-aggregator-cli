# News Aggregator CLI

A Python command-line application that collects news headlines from multiple RSS feeds, filters articles, stores them locally, and exports news data.

## Features

- Fetch news from multiple RSS sources.
- Filter articles by source, keyword, and date.
- Store articles in SQLite and JSON.
- Avoid duplicate articles.
- Export news to CSV, JSON, and Excel.
- Simple command-line interface.

## Requirements

- Python 3.10 or newer
- Internet connection

## Installation

Install the required libraries:

```bash
pip install -r requirements.txt
```

## Usage

Fetch news:

```bash
python main.py fetch
```

List saved articles:

```bash
python main.py list
```

Filter by keyword:

```bash
python main.py list --keyword technology
```

Filter by source:

```bash
python main.py list --source BBC
```

Export articles:

```bash
python main.py export --output news.csv
```

Export to Excel:

```bash
python main.py export --output news.xlsx
```

## Technologies Used

- Python
- RSS feeds
- SQLite
- JSON
- CSV
- OpenPyXL

## Author

Add your name here.
