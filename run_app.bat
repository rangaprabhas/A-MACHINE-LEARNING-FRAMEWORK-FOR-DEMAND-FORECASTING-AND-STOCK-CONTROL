@echo off
title Demand Forecasting and Stock Control Framework
color 0b
echo Starting Streamlit local server...
start "" http://localhost:8501
streamlit run ML_project_main.py --server.port 8501
pause
