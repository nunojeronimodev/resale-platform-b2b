# 🛒 Resale Platform B2B — supplier price aggregator

A Flask + pandas web app that merges the price lists of several technology suppliers into one
catalogue, finds the same product across suppliers, keeps the lowest price, applies sales margins
and exports Excel price lists ready to send to customers.

> [Flask](https://flask.palletsprojects.com/) is a lightweight Python web framework, used here for the backend and the web interface.

> [pandas](https://pandas.pydata.org/) is a fast and flexible data analysis library, used here to read, clean and merge the supplier price lists.

## 🔍 Demo

<!-- Record a short GIF of the app, save it as docs/demo.gif and uncomment the line below -->
<!-- ![View Demo](docs/demo.gif) -->

_Demo coming soon._ 🎬

## ✨ Features

- 📥 Import CSV and Excel files with automatic column detection
- 🔄 Automatic sync by URL every 24 hours, or manual upload
- 🔗 Smart deduplication across suppliers by EAN, part number and product name
- 💰 Reseller and VIP margins, with Excel export per category
- 🛡️ Automatic audit of the final catalogue (duplicates, accessories, suspicious prices)
- 📈 Price history tool that tracks price changes between supplier updates

## ⚙️ How it works

[How the application works]

<img width="797" height="552" alt="DIAGRAMA DE FLUXOS DE DADOS" src="https://github.com/user-attachments/assets/e9698918-8ca9-4ecc-b16c-919223fbc7d7" />

Each supplier's price list is imported, mapped to a common format and filtered. Then the app
detects the same product across suppliers, using exact keys first (EAN, part number), then a set
of guards (model, storage, 4G/5G, colour) and finally name similarity, so that a *Galaxy S24* is
never mixed up with a *Galaxy S25*. 😅 Only the lowest price is kept, margins are applied and the
result is audited before the Excel files are generated.

## 📋 Requirements

You may use Windows, macOS or Linux as your operating system, with the following tools:

- [Git](https://git-scm.com/)
- [Python 3.10+](https://www.python.org/downloads/)

## 🚀 Technologies

This project was built using the following technologies:

- [Python](https://www.python.org/)
- [Flask](https://flask.palletsprojects.com/)
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [pytest](https://docs.pytest.org/)
- HTML, CSS and JavaScript

## 💻 Installation

To clone this project, run the following command:

```bash
git clone https://github.com/nunojeronimodev/resale-platform-B2B.git
cd resale-platform-B2B/catalogo
```

Create a virtual environment and install the dependencies:

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
source .venv/bin/activate     # macOS / Linux
pip install -r requirements.txt
```

Then start the application:

```bash
python app.py
```

Open **http://localhost:5050** in your browser and add your suppliers. 🎉

> 💡 On Windows you can simply double-click `catalogo/INICIAR.bat`.

> 🇵🇹 Note: the user interface is in Portuguese, as the app was built for the Portuguese market.

## 🌐 Hosting

The app is designed to run on **localhost** (`127.0.0.1:5050`), but you can easily adapt it to
another host or your own domain. The host and port are set in the last line of
`catalogo/app.py`:

```python
app.run(host="127.0.0.1", debug=False, port=5050, threaded=True)
```

Change them to suit your setup, or deploy the app on any server that supports Python, using a
production WSGI server such as [Gunicorn](https://gunicorn.org/) or
[Waitress](https://docs.pylonsproject.org/projects/waitress/) behind a reverse proxy
(e.g. Nginx) with your domain.

> ⚠️ The app has no built-in login. Before making it available on a network or on the internet,
> add an authentication layer (for example, a reverse proxy with login and HTTPS, or a VPN such as
> [Tailscale](https://tailscale.com/)), so that only you can see your suppliers, prices and margins.

## 🧪 Tests

```bash
pip install -r requirements-dev.txt
pytest
```

## 👍 Contributing

Contributions, ideas and feedback are always welcome! Feel free to open an
[issue](https://github.com/nunojeronimodev/resale-platform-B2B/issues) to report a problem or
suggest an improvement.

## 📄 Licensing

This repository is MIT licensed, as found in the [LICENSE](LICENSE) file.

---

Made with ☕ by [Nuno Jerónimo](https://github.com/nunojeronimodev)
