[app]
title = Cone Calculator
package.name = conecalculator
package.domain = org.ken
source.dir = .
source.main = main.py
source.include_exts = py,png,jpg,kv,atlas,json,txt
version = 1.2.0

# Added pillow and et_xmlfile which are often needed by openpyxl
requirements = python3==3.10.14,kivy==2.3.0,openpyxl,et_xmlfile,pillow,sdl2,sdl2_image,sdl2_mixer,sdl2_ttf

android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.minapi = 21
# Downgraded from 36 to 34 for stability with Buildozer 1.5.0
android.api = 34
android.ndk = 25c
android.archs = arm64-v8a
android.accept_sdk_license = True
android.enable_androidx = True
android.gradle_dependencies = androidx.appcompat:appcompat:1.6.1

orientation = portrait
fullscreen = 0
log_level = 2
warn_on_root = 1

[buildozer]
build_dir = .buildozer
cache_dir = .buildozer_cache
