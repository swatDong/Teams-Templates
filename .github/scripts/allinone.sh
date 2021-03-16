#!/usr/bin/env bash
set -x

# to add input checking.
RAW_TAG=$1

# get folder path to zip.
if [[ ${RAW_TAG} == *"tab.JavaScript.default"* ]];then
    cd ./tab/js/default
elif [[ ${RAW_TAG} == *"tab.JavaScript.with-function"* ]];then
    cd ./tab/js/with-function
elif [[ ${RAW_TAG} == *"bot.JavaScript.default"* ]];then
    cd ./bot/js
elif [[ ${RAW_TAG} == *"bot.TypeScript.default"* ]];then
    cd ./bot/ts
else # botCsDefault
    cd ./bot/csharp
fi

zip -r ../../${RAW_TAG}.zip .
cd ../..
