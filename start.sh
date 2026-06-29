#!/bin/bash
mkdir -p /data/uploads
python init_db.py
exec gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 60
