#!/usr/bin/env bash
set -x

# to add input checking.
RAW_TAG=$1

# get folder path to zip.

if [[ ${RAW_TAG} == *"botJsDefault"* ]];then
    cd ./bot/js
elif [[ ${RAW_TAG} == *"botTsDefault"* ]];then
    cd ./bot/ts
else # botCsDefault
    cd ./bot/csharp
fi

zip ../../${RAW_TAG}.zip .
cd ../..
