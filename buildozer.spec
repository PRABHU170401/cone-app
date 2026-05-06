name: Build Android APK

on:
  push:
    branches: [ "main", "master" ]
  pull_request:
    branches: [ "main", "master" ]
  workflow_dispatch:

jobs:
  build-apk:
    name: Build APK with Buildozer
    runs-on: ubuntu-22.04

    steps:
      # ── 1. Check out the repository ──────────────────────────────────────────
      - name: Checkout code
        uses: actions/checkout@v4

      # ── 2. Rename main Python file to main.py ────────────────────────────────
      - name: Rename entry point to main.py
        run: |
          # Find the app Python file and rename it to main.py
          # (Buildozer requires the entry point to be named main.py)
          if [ ! -f main.py ]; then
            PY_FILE=$(ls cone_calculator*.py 2>/dev/null | head -1)
            if [ -n "$PY_FILE" ]; then
              cp "$PY_FILE" main.py
              echo "Renamed $PY_FILE → main.py"
            else
              echo "ERROR: No cone_calculator*.py file found!"
              ls -la *.py || true
              exit 1
            fi
          else
            echo "main.py already exists, skipping rename."
          fi

      # ── 3. Set up Python ─────────────────────────────────────────────────────
      - name: Set up Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      # ── 4. Install system dependencies for Buildozer / Android SDK ───────────
      - name: Install system dependencies
        run: |
          sudo apt-get update -qq
          sudo apt-get install -y --no-install-recommends \
            git zip unzip python3-pip \
            autoconf libtool pkg-config \
            zlib1g-dev libncurses5-dev libncursesw5-dev \
            libtinfo5 cmake libffi-dev libssl-dev \
            ccache openjdk-17-jdk

      # ── 5. Cache Buildozer downloads (.buildozer_cache) ──────────────────────
      - name: Cache Buildozer download cache
        uses: actions/cache@v4
        with:
          path: .buildozer_cache
          key: buildozer-cache-${{ runner.os }}-${{ hashFiles('buildozer.spec') }}
          restore-keys: |
            buildozer-cache-${{ runner.os }}-

      # ── 6. Cache the .buildozer build directory (speeds up rebuilds) ─────────
      - name: Cache Buildozer build directory
        uses: actions/cache@v4
        with:
          path: .buildozer
          key: buildozer-build-${{ runner.os }}-${{ hashFiles('buildozer.spec') }}-${{ hashFiles('**/*.py') }}
          restore-keys: |
            buildozer-build-${{ runner.os }}-

      # ── 7. Install Buildozer and Cython ──────────────────────────────────────
      - name: Install Buildozer
        run: |
          pip install --upgrade pip
          pip install buildozer cython

      # ── 8. Build debug APK ───────────────────────────────────────────────────
      - name: Build APK (debug)
        env:
          ANDROID_SDK_ROOT: /home/runner/.buildozer/android/platform/android-sdk
        run: |
          buildozer -v android debug 2>&1 | tee build.log

      # ── 9. Show bin/ contents (debug: helps confirm APK was created) ─────────
      - name: List bin directory
        if: always()
        run: |
          echo "=== bin/ contents ==="
          ls -lh bin/ 2>/dev/null || echo "bin/ directory does not exist!"
          echo "=== APK search ==="
          find . -name "*.apk" 2>/dev/null || echo "No .apk files found anywhere."

      # ── 10. Upload the APK as a build artifact ───────────────────────────────
      - name: Upload APK artifact
        uses: actions/upload-artifact@v4
        with:
          name: cone-calculator-debug-apk
          path: bin/*.apk
          if-no-files-found: error
          retention-days: 30

      # ── 11. Upload build log for debugging ───────────────────────────────────
      - name: Upload build log
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: build-log
          path: build.log
          retention-days: 7
