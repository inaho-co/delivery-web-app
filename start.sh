#!/bin/sh
export PYTHONPATH=/volume1/homes/nastom/.local/lib/python3.8/site-packages
set -a
. /volume1/webapp/.env
set +a
cd /volume1/webapp
nohup python3 app.py >> /tmp/app.log 2>&1 &
echo "Started PID: $!"
