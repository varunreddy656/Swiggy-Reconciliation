#!/bin/bash
python3 -m pip install --upgrade pip
pip install -r requirements.txt
exec python3 -m gunicorn app:app
