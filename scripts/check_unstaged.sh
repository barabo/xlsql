#!/bin/sh
if git diff --exit-code --quiet
then
    exit 0
else
    echo "There are unstaged changes."
    exit 1
fi
