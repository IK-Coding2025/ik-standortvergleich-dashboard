"""Führt die komplette Pipeline aus: Datenabruf (fetch) -> Merge.

    python run_pipeline.py

Danach das Dashboard starten mit:
    streamlit run dashboard.py
"""

import fetch_data
import merge_data

if __name__ == "__main__":
    fetch_data.main()
    merge_data.main()
    print("Pipeline abgeschlossen. Dashboard starten mit: "
          "streamlit run dashboard.py")
