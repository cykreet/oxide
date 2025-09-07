#!/bin/bash

release=$([ "$1" == "--release" ] && echo "--release" || echo "")
cargo build --target x86_64-pc-windows-gnu $release