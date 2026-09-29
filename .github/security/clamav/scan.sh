#!/bin/sh
# Runs inside the ClamAV image: fetch today's signatures, then scan /scan.
# Everything goes to /out, and the status file says how the scan ended, so a
# failed update or a scanner error reaches the security report as an error.
set -u
if ! freshclam --stdout --foreground > /out/freshclam.log 2>&1; then
    echo "freshclam could not update the signatures; see freshclam.log" > /out/clamav.err
    echo 2 > /out/clamav.status
    exit 0
fi
clamscan --version > /out/clamav-version.txt
# --alert-macros and --alert-encrypted widen the net past known malware to any
# Office file with VBA and any encrypted archive or document.
clamscan -r --infected --no-summary --alert-macros=yes --alert-encrypted=yes \
    /scan > /out/clamav.log 2> /out/clamav.err
echo $? > /out/clamav.status
