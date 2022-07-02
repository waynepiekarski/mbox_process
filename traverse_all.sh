#!/bin/bash

BASEDIR="$1"
if [[ "${BASEDIR}" == "" ]]; then
    echo "Base directory not specified"
    exit 1
fi
OUTDIR="$2"
if [[ "${OUTDIR}" == "" ]]; then
    OUTDIR="$(dirname $0)/OUTPUT"
fi
rm -rf "${OUTDIR}"
mkdir "${OUTDIR}"

# Does not support spaces in the directory names
time for each in $( find -L "${BASEDIR}" -type f -printf "%P\n" ); do
    IN="${BASEDIR}/${each}"
    OUT="${OUTDIR}/$(echo ${each} | sed "s~/~-~g")"
    if [[ $IN == *.gz ]]; then
	echo "DECOMPRESSING [${IN}]"
	gunzip --to-stdout "${IN}" >> "tmp-unzipped.mbox"
	IN="tmp-unzipped.mbox"
    fi
    echo "PROCESSING input [${IN}] output [${OUT}]"
    python3 "$(dirname $0)"/mbox_process.py "${IN}" "${OUT}" &> "${OUT}.log"
    RESULT="$?"
    rm -f "tmp-unzipped.mbox" &> /dev/null
    if [[ "${RESULT}" != "0" ]]; then
	echo "ERROR $each"
	touch "${OUT}.fail"
	tail -1 "${OUT}.log"
    fi
done
