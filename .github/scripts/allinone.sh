#!/usr/bin/env bash
set -x

# to add input checking.
RAW_TAG=$1

# get folder path to zip.
# tab
if [[ ${RAW_TAG} == *"tab.JavaScript.default"* ]];then
    cd ./tab/js/default
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"tab.JavaScript.with-function"* ]];then
    cd ./tab/js/with-function
    zip -r ../../../${RAW_TAG}.zip .
# function
elif [[ ${RAW_TAG} == *"function-base.JavaScript.default"* ]];then
    cd ./function-base/js/default
    zip -r ../../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"function-triggers.JavaScript.HTTPTrigger"* ]];then
    cd ./function-triggers/js/HTTPTrigger
    zip -r ../../../${RAW_TAG}.zip .
# bot
elif [[ ${RAW_TAG} == *"bot.JavaScript.default"* ]];then
    cd ./bot/js
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"bot.TypeScript.default"* ]];then
    cd ./bot/ts
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"bot-msgext.JavaScript.default"* ]];then
    cd ./bot-msgext/js
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"bot-msgext.TypeScript.default"* ]];then
    cd ./bot-msgext/ts
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"msgext.JavaScript.default"* ]];then
    cd ./msgext/js
    zip -r ../../${RAW_TAG}.zip .
elif [[ ${RAW_TAG} == *"msgext.TypeScript.default"* ]];then
    cd ./msgext/ts
    zip -r ../../${RAW_TAG}.zip .
else 
    echo "Unknown tag ${RAW_TAG}"
fi

cd ../..
