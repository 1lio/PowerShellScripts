echo off

ipconfig /release
arp -d *
nbtstat -R
ipconfig /flushdns
ipconfig /renew

net stop dnscache
net start dnscache
