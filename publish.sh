#!/usr/bin/env bash
set -euo pipefail

# Build NativeAOT binaries for all supported platforms.
# Requires .NET 10 SDK and Rust toolchain.
#
# Usage:
#   ./publish.sh              # Build for current platform
#   ./publish.sh all          # Build for all platforms (cross-compile)
#   ./publish.sh macos-arm64  # Build for specific target

SERVER_PROJECT="src/DocxMcp/DocxMcp.csproj"
CLI_PROJECT="src/DocxMcp.Cli/DocxMcp.Cli.csproj"
STORAGE_CRATE="crates/docx-storage-local"
OUTPUT_DIR="dist"
CONFIG="Release"

declare -A TARGETS=(
    ["macos-arm64"]="osx-arm64"
    ["macos-x64"]="osx-x64"
    ["linux-x64"]="linux-x64"
    ["linux-arm64"]="linux-arm64"
    ["windows-x64"]="win-x64"
    ["windows-arm64"]="win-arm64"
)

# Rust target triples for cross-compilation
declare -A RUST_TARGETS=(
    ["macos-arm64"]="aarch64-apple-darwin"
    ["macos-x64"]="x86_64-apple-darwin"
    ["linux-x64"]="x86_64-unknown-linux-gnu"
    ["linux-arm64"]="aarch64-unknown-linux-gnu"
    ["windows-x64"]="x86_64-pc-windows-msvc"
    ["windows-arm64"]="aarch64-pc-windows-msvc"
)

publish_project() {
    local project="$1"
    local binary_name="$2"
    local rid="$3"
    local out="$4"

    dotnet publish "$project" \
        --configuration "$CONFIG" \
        --runtime "$rid" \
        --self-contained true \
        --output "$out" \
        -p:PublishAot=true \
        -p:OptimizationPreference=Size

    local binary
    if [[ "$out" == *windows* ]]; then
        binary="$out/${binary_name}.exe"
    else
        binary="$out/$binary_name"
    fi

    if [[ -f "$binary" ]]; then
        local size
        size=$(du -sh "$binary" | cut -f1)
        echo "    Built: $binary ($size)"
    else
        echo "    WARNING: Binary not found at $binary"
    fi
}

publish_rust_storage() {
    local name="$1"
    local out="$2"
    local rust_target="${RUST_TARGETS[$name]}"
    local current_target

    # Detect current Rust target
    local arch
    arch="$(uname -m)"
    case "$(uname -s)-$arch" in
        Darwin-arm64) current_target="aarch64-apple-darwin" ;;
        Darwin-x86_64) current_target="x86_64-apple-darwin" ;;
        Linux-x86_64) current_target="x86_64-unknown-linux-gnu" ;;
        Linux-aarch64) current_target="aarch64-unknown-linux-gnu" ;;
        *) current_target="" ;;
    esac

    local binary_name="docx-storage-local"
    [[ "$name" == windows-* ]] && binary_name="docx-storage-local.exe"

    if [[ "$rust_target" == "$current_target" ]]; then
        # Native build
        echo "    Building Rust storage server (native)..."
        cargo build --release --package docx-storage-local
        cp "target/release/$binary_name" "$out/" 2>/dev/null || \
            cp "target/release/docx-storage-local" "$out/$binary_name"
    else
        # Cross-compile (requires target installed)
        if rustup target list --installed | grep -q "$rust_target"; then
            echo "    Building Rust storage server (cross: $rust_target)..."
            cargo build --release --package docx-storage-local --target "$rust_target"
            cp "target/$rust_target/release/$binary_name" "$out/" 2>/dev/null || \
                cp "target/$rust_target/release/docx-storage-local" "$out/$binary_name"
        else
            echo "    SKIP: Rust target $rust_target not installed (run: rustup target add $rust_target)"
            return 0
        fi
    fi

    if [[ -f "$out/$binary_name" ]]; then
        local size
        size=$(du -sh "$out/$binary_name" | cut -f1)
        echo "    Built: $out/$binary_name ($size)"
    fi
}

publish_target() {
    local name="$1"
    local rid="${TARGETS[$name]}"
    local out="$OUTPUT_DIR/$name"

    mkdir -p "$out"

    # On macOS, NativeAOT needs Homebrew libraries (openssl, brotli, etc.)
    if [[ "$(uname -s)" == "Darwin" ]]; then
        export LIBRARY_PATH="/opt/homebrew/lib:${LIBRARY_PATH:-}"
    fi

    echo "==> Publishing docx-storage-local ($name)..."
    publish_rust_storage "$name" "$out"

    echo "==> Publishing docx-mcp ($name / $rid)..."
    publish_project "$SERVER_PROJECT" "docx-mcp" "$rid" "$out"

    echo "==> Publishing docx-cli ($name / $rid)..."
    publish_project "$CLI_PROJECT" "docx-cli" "$rid" "$out"
}

publish_rust_only() {
    local rid_name="$1"
    local out="$OUTPUT_DIR/$rid_name"
    mkdir -p "$out"

    echo "==> Publishing docx-storage-local ($rid_name)..."
    publish_rust_storage "$rid_name" "$out"
}

detect_current_platform() {
    local arch
    arch="$(uname -m)"
    case "$(uname -s)-$arch" in
        Darwin-arm64) echo "macos-arm64" ;;
        Darwin-x86_64) echo "macos-x64" ;;
        Linux-x86_64) echo "linux-x64" ;;
        Linux-aarch64) echo "linux-arm64" ;;
        *) echo ""; return 1 ;;
    esac
}

main() {
    local target="${1:-current}"

    echo "docx-mcp NativeAOT publisher"
    echo "=============================="

    if [[ "$target" == "all" ]]; then
        for name in "${!TARGETS[@]}"; do
            publish_target "$name"
        done
    elif [[ "$target" == "rust" ]]; then
        # Build only Rust storage server for current platform
        local rid_name
        rid_name=$(detect_current_platform) || { echo "Unsupported platform"; exit 1; }
        publish_rust_only "$rid_name"
    elif [[ "$target" == "current" ]]; then
        # Detect current platform
        local rid_name
        rid_name=$(detect_current_platform) || { echo "Unsupported platform: $(uname -s)-$(uname -m)"; exit 1; }
        publish_target "$rid_name"
    elif [[ -n "${TARGETS[$target]+x}" ]]; then
        publish_target "$target"
    else
        echo "Unknown target: $target"
        echo "Available: ${!TARGETS[*]} all current rust"
        exit 1
    fi

    echo ""
    echo "Done. Binaries are in $OUTPUT_DIR/"
}

main "$@"
