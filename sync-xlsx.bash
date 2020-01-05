#!/usr/bin/env bash

SRC=/home/pompei/trans
TO=/home/pompei/IdeaProjects/greetgo.msoffice/test_src/kz/greetgo/msoffice/xlsx/reader/xlsx
ARCH=/home/pompei/IdeaProjects/greetgo.msoffice/tmp

for f in $(find ${SRC} -type f -name "*.xlsx" -and -not -name "~*") ; do

  cp -f ${f} ${TO}/

  name=$(basename ${f})

  rm     -rf  ${ARCH}/${name}
  mkdir  -p   ${ARCH}/${name}
  cd          ${ARCH}/${name}
  unzip       ${f}

done
