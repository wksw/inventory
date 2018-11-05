#!/bin/bash

ROOTDIR=$(cd $(dirname $0);pwd)

cat << EOF > $ROOTDIR/crontab
30 23 * * * /bin/bash $ROOTDIR/get_parse.sh >> $ROOTDIR/parse.log
EOF

crontab $ROOTDIR/crontab
rm -rf $ROOTDIR/crontab

