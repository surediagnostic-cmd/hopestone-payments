#!/bin/bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
curl -o "$SCRIPT_DIR/weekly_cases.csv" \
  'https://d33hv1py9n6skn.cloudfront.net/uploads/production/data_export_files/exported_file/12083/cases-2026-03-30-38020260330-2-y2px54.csv' \
  && echo "SUCCESS: CSV saved to $SCRIPT_DIR/weekly_cases.csv" \
  || echo "FAILED"
sleep 2
