#!/bin/bash
cd '/c/Users/User/Desktop/Data logging'
git add .
git commit -m 'Initial commit'
git branch -M main
git remote add origin https://github.com/tawornp07-lab/bioprocess-data-logging.git
git push -u origin main
