[app]

# App title shown on Android launcher
title = Cone Calculator

# Package name (must be unique, lowercase, no spaces)
package.name = conecalculator

# Package domain (reverse domain notation)
package.domain = org.ken

# Source directory (relative to buildozer.spec)
source.dir = .

# Main entry point
source.include_exts = py,png,jpg,kv,atlas,json,txt

# App version
version = 1.2.0

# Requirements — all packages the app needs
requirements = python3,kivy==2.3.0,openpyxl

# Android permissions
android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,INTERNET

# Minimum and target Android API
android.minapi = 21
android.api = 33
android.ndk = 25b

# Android architecture(s) — arm64-v8a covers most modern phones
android.archs = arm64-v8a

# Accept Android SDK license automatically in CI
android.accept_sdk_license = True

# Orientation
orientation = portrait

# Fullscreen (0 = show status bar, 1 = hide it)
fullscreen = 0

# Icon (place a 512x512 icon.png next to buildozer.spec to use it)
# icon.filename = %(source.dir)s/icon.png

# Presplash (place a presplash.png next to buildozer.spec to use it)
# presplash.filename = %(source.dir)s/presplash.png

# Log level: 0 = error only, 1 = info, 2 = debug
log_level = 2

# Warn when buildozer is newer than what the spec was written for
warn_on_root = 1

[buildozer]
# Build directory (relative)
build_dir = .buildozer

# Cache downloaded files here
cache_dir = .buildozer_cache