# Flights Data Quality — build targets
# Usage: make all [DQ_DATA=path/to/flights.csv] [DQ_REPORTS=path/to/out]
PYTHON ?= python3
export DQ_DATA
export DQ_REPORTS

.PHONY: all check docx xlsx install clean help

all: check docx xlsx ## Run the checks and build both reports

install: ## Install Python dependencies
	$(PYTHON) -m pip install -r requirements.txt

check: ## Print pass/fail counts for rules R01–R18
	cd src && $(PYTHON) analysis.py

docx: ## Build the Word report
	cd src && $(PYTHON) build_docx.py

xlsx: ## Build the Excel workbook
	cd src && $(PYTHON) build_xlsx.py

clean: ## Remove generated reports and caches
	rm -f reports/*.docx reports/*.xlsx
	rm -rf src/__pycache__

help: ## List targets
	@grep -E '^[a-z]+:.*##' $(MAKEFILE_LIST) | awk -F':.*## ' '{printf "  %-8s %s\n", $$1, $$2}'
