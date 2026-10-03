#!/usr/bin/env bash
# BTA OSG External Intake Portal: Oracle Cloud Ubuntu Setup Script
# Target OS: Ubuntu 22.04 LTS or 24.04 LTS (ARM64 / x86_64)

set -euo pipefail

echo ">>> [1/7] Updating apt packages..."
export DEBIAN_FRONTEND=noninteractive
sudo apt-get update -y
sudo apt-get upgrade -y

echo ">>> [2/7] Installing Nginx, PHP 8.2 FPM, MySQL 8, and dependencies..."
sudo apt-get install -y \
    nginx \
    php8.2-fpm \
    php8.2-mysql \
    php8.2-curl \
    php8.2-mbstring \
    mysql-server \
    certbot \
    python3-certbot-nginx \
    curl \
    git \
    ufw \
    iptables-persistent \
    rclone

echo ">>> [3/7] Configuring Host Firewall (Dual-Firewall Requirement)..."
# Oracle Cloud images ship with restrictive iptables rules by default.
# Open HTTP (80), HTTPS (443), and SSH (22) in both iptables and ufw.
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT || true
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 443 -j ACCEPT || true
sudo netfilter-persistent save || true

sudo ufw allow 22/tcp comment 'SSH'
sudo ufw allow 80/tcp comment 'HTTP'
sudo ufw allow 443/tcp comment 'HTTPS'
sudo ufw --force enable

echo ">>> [4/7] Securing MySQL to 127.0.0.1 (Localhost Only)..."
MYSQL_CNF="/etc/mysql/mysql.conf.d/mysqld.cnf"
if [ -f "$MYSQL_CNF" ]; then
    sudo sed -i 's/^bind-address.*/bind-address = 127.0.0.1/' "$MYSQL_CNF"
    sudo systemctl restart mysql
fi

echo ">>> [5/7] Preparing Web Directory and Permissions..."
WEB_ROOT="/var/www/osg-portal"
sudo mkdir -p "$WEB_ROOT"
sudo chown -R www-data:www-data "$WEB_ROOT"
sudo chmod -R 755 "$WEB_ROOT"

echo ">>> [6/7] Enabling Nginx and PHP-FPM Services..."
sudo systemctl enable nginx
sudo systemctl start nginx
sudo systemctl enable php8.2-fpm
sudo systemctl start php8.2-fpm
sudo systemctl enable mysql
sudo systemctl start mysql

echo ">>> [7/7] Installation Complete!"
echo "--- CRITICAL ORACLE CLOUD REMINDERS ---"
echo " 1. Open Ports 80 and 443 in the ORACLE CLOUD CONSOLE:"
echo "    Networking -> Virtual Cloud Networks -> VCN -> Default Security List"
echo "    Add Ingress Rules for 0.0.0.0/0 on Destination Port 80 and 443."
echo " 2. Import database schema: mysql -u root < portal/db/schema.sql"
echo " 3. Deploy code to: $WEB_ROOT"
echo " 4. Obtain SSL cert: sudo certbot --nginx -d your-portal-domain.com"
