[app]

title = Cone Calculator
package.name = conecalculator
package.domain = org.ken
source.dir = .
source.main = main.py
source.include_exts = py,png,jpg,kv,atlas,json,txt
version = 1.2.0

requirements = python3==3.10.14,kivy==2.3.0,openpyxl,sdl2,sdl2_image,sdl2_mixer,sdl2_ttf

android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,INTERNET
android.minapi = 21
android.api = 33
android.ndk = 25b
android.archs = arm64-v8a
android.accept_sdk_license = True
android.enable_androidx = True
android.gradle_dependencies = androidx.appcompat:appcompat:1.4.0

orientation = portrait
fullscreen = 0
log_level = 2
warn_on_root = 1

[buildozer]
build_dir = .buildozer
cache_dir = .buildozer_cache