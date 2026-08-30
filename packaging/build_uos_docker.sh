#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="$(sed -n 's/^__version__ = "\([^"]*\)"/\1/p' "$ROOT/wfp_version.py")"
ARCH="${1:-aarch64}"

if [[ -z "$VERSION" ]]; then
  echo "Unable to read version from wfp_version.py" >&2
  exit 1
fi

case "$ARCH" in
  aarch64)
    PLATFORM="linux/arm64"
    TARGETARCH="arm64"
    ;;
  x86_64)
    PLATFORM="linux/amd64"
    TARGETARCH="amd64"
    ;;
  *)
    echo "Usage: $0 [aarch64|x86_64]" >&2
    exit 2
    ;;
esac

IMAGE="${WFP_UOS_IMAGE:-word-formatter-pro-uos-${TARGETARCH}:${VERSION}}"
RAW="release/Word-Formatter-Pro.v${VERSION}.UOS-V20.${ARCH}"
APPIMAGE="${RAW}.AppImage"
BUILD_ARGS=(--build-arg "TARGETARCH=${TARGETARCH}")

# These overrides are useful on an offline Windows build host or behind a
# restricted proxy. The defaults remain macrosan/uos and conda-forge.
if [[ -n "${WFP_UOS_BASE_IMAGE:-}" ]]; then
  BUILD_ARGS+=(--build-arg "UOS_BASE_IMAGE=${WFP_UOS_BASE_IMAGE}")
fi
if [[ -n "${WFP_CONDA_CHANNEL:-}" ]]; then
  BUILD_ARGS+=(--build-arg "CONDA_CHANNEL=${WFP_CONDA_CHANNEL}")
fi

docker build \
  --platform "$PLATFORM" \
  "${BUILD_ARGS[@]}" \
  --tag "$IMAGE" \
  --file "$ROOT/packaging/Dockerfile.uos" \
  "$ROOT"

docker run --rm \
  --platform "$PLATFORM" \
  --volume "$ROOT:/workspace" \
  --workdir /workspace \
  "$IMAGE" \
  python packaging/build_release.py uos \
    --arch "$ARCH" \
    --reuse-venv \
    --also-raw \
    --appimagetool /usr/local/bin/appimagetool

# The raw executable is always emitted because it works when UOS has no FUSE.
# APPIMAGE_EXTRACT_AND_RUN also lets CI validate the AppImage without FUSE.
docker run --rm \
  --platform "$PLATFORM" \
  --volume "$ROOT:/workspace:ro" \
  --workdir /workspace \
  --env APPIMAGE_EXTRACT_AND_RUN=1 \
  "$IMAGE" \
  bash -lc "
    set -euo pipefail
    test \"\$(uname -m)\" = '$ARCH'
    file '$RAW' '$APPIMAGE'
    test \"\$('${RAW}' --version)\" = '$VERSION'
    test \"\$('${APPIMAGE}' --version)\" = '$VERSION'
    '${RAW}' --test
    sha256sum '$RAW' '$APPIMAGE'
  "

echo "UOS ${ARCH} artifacts:"
echo "  $ROOT/$RAW"
echo "  $ROOT/$APPIMAGE"
