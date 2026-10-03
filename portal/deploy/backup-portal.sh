#!/usr/bin/env bash
# BTA OSG External Intake Portal: Automated Daily Backup Script
# Bangsamoro Transition Authority - Office of the Secretary-General
# Creates gzipped MySQL dump and portal code archive in /var/backups/osg/
# Retains last 30 daily backups.

set -euo pipefail

BACKUP_DIR="/var/backups/osg"
DATE_TAG=$(date +"%Y%m%d_%H%M%S")
DB_NAME="osg"
DB_USER="portal"
WEB_ROOT="/var/www/osg-portal"
# mysqldump must authenticate explicitly: it never inherits ambient credentials, and
# relying on a defaults-file silently dumps nothing for www-data. Set the portal user's
# password in /etc/osg-portal-backup.cnf (chmod 600, owned by root):
#   [client]
#   user=portal
#   password=...
CNF_FILE="/etc/osg-portal-backup.cnf"

mkdir -p "${BACKUP_DIR}"
chmod 700 "${BACKUP_DIR}"

if [ ! -r "${CNF_FILE}" ]; then
    echo "[$(date -u +"%Y-%m-%d %H:%M:%S UTC")] FATAL: credential file ${CNF_FILE} missing or unreadable; aborting without a dump." >&2
    exit 1
fi

echo "[$(date -u +"%Y-%m-%d %H:%M:%S UTC")] Starting BTA OSG portal backup..."

# 1. Backup MySQL Database
SQL_FILE="${BACKUP_DIR}/osg_db_${DATE_TAG}.sql.gz"
echo "Dumping MySQL database '${DB_NAME}'..."
mysqldump --defaults-extra-file="${CNF_FILE}" --single-transaction --quick --routines --triggers "${DB_NAME}" | gzip -9 > "${SQL_FILE}"
chmod 600 "${SQL_FILE}"
echo "Database dump complete: ${SQL_FILE} ($(du -h "${SQL_FILE}" | cut -f1))"

# 2. Backup Portal Code and Config
TAR_FILE="${BACKUP_DIR}/osg_web_${DATE_TAG}.tar.gz"
echo "Archiving web files from '${WEB_ROOT}'..."
tar -czf "${TAR_FILE}" -C "/var/www" "osg-portal"
chmod 600 "${TAR_FILE}"
echo "Web files archive complete: ${TAR_FILE} ($(du -h "${TAR_FILE}" | cut -f1))"

# 3. Retention policy: Remove backups older than 30 days
echo "Pruning backups older than 30 days..."
find "${BACKUP_DIR}" -type f -name "osg_*" -mtime +30 -delete

echo "[$(date -u +"%Y-%m-%d %H:%M:%S UTC")] Backup successfully completed."
