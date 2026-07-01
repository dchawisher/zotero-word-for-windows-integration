#!/bin/bash

CWD="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
TEMPLATE_DIR="$CWD/Citate.dotm"
BIN_DIR="$TEMPLATE_DIR/word/vbaProject.bin"

# Citate.dotm
echo 'Unpacking Citate.dotm...'

cd "$( dirname "${BASH_SOURCE[0]}" )"
rm -rf $TEMPLATE_DIR*

# Other than vbaProject.bin, all files are XML
unzip -q ../../install/Citate.dotm -d $TEMPLATE_DIR

# Pretty-print XML files
find Citate.dotm/ -type f \( -iname '*.xml' -o -iname '*.rels' \) -exec  xmllint --output '{}' --format '{}' \;

# Extract vbaProject.bin
rm $BIN_DIR
mkdir $BIN_DIR
python ../tools/officeparser/officeparser.py -l ERROR -o $BIN_DIR \
       --extract-macros ../../install/Citate.dotm

# Remove unnecessary files
rm -rf $TEMPLATE_DIR/_rels $TEMPLATE_DIR/docProps $TEMPLATE_DIR/word/_rels \
	$TEMPLATE_DIR/word/theme $TEMPLATE_DIR/word/*.xml $TEMPLATE_DIR/customUI/_rels \
	"$TEMPLATE_DIR/[Content_Types].xml"