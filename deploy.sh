#!/bin/bash
git pull origin main
sudo cp public/index.html /var/www/html/index.html
sudo systemctl reload nginx