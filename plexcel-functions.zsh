# plexcel shell functions
# Add to your .zshrc:
#   export PLEXCEL_DIR="/path/to/sheets-styling"
#   source "$PLEXCEL_DIR/plexcel-functions.zsh"

: ${PLEXCEL_DIR:="$(cd "$(dirname "$0")" && pwd)"}

# Convert CSV to XLSX
csv-to-xlsx() {
    if [[ $# -lt 2 ]]; then
        echo "Usage: csv-to-xlsx output.xlsx workbook.yaml"
        echo "       csv-to-xlsx output.xlsx file.csv:Sheet1 [file.csv:Sheet2 ...]"
        return 1
    fi
    "$PLEXCEL_DIR/venv/bin/python" "$PLEXCEL_DIR/csv_to_xlsx.py" "$@"
}

# Edit shared styles
plexcel-styles() {
    ${EDITOR:-vim} "$PLEXCEL_DIR/styles.yaml"
}

# Create starter workbook.yaml in current dir
plexcel-init() {
    local f="${1:-workbook.yaml}"
    cat > "$f" <<'EOF'
sheets:
  - csv: data.csv
    name: Sheet1
EOF
    echo "Created $f"
}
