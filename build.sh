#!/usr/bin/env bash
set -euo pipefail
version="${1:-${PLENARO_VERSION:-}}"
if [[ -z "$version" ]]; then
  tag="$(git describe --tags --exact-match 2>/dev/null || true)"
  [[ "$tag" =~ ^[vV]([0-9]+\.[0-9]+\.[0-9]+([.-][0-9A-Za-z.-]+)?)$ ]] && version="${BASH_REMATCH[1]}"
fi
version="${version:-2.1.0-dev}"
if [[ "$version" =~ ^[vV](.+)$ ]]; then version="${BASH_REMATCH[1]}"; fi
dotnet publish TaskTool.Wpf.csproj -c Release -p:Version="$version" -p:InformationalVersion="$version"
