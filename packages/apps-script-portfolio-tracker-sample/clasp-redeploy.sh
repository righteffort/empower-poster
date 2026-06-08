#!/usr/bin/env bash
set -euo pipefail

clasp push -f

# Find the single non-HEAD deployment
deployments_output=$(clasp list-deployments)
deployment_id=$(echo "$deployments_output" |
		    grep -e '^-' | 
		    awk '$3 != "@HEAD" {print $2}')
[[ $(echo "$deployment_id" | wc -w) -eq 1 ]] || { 
    echo "FAILED. Script only works with exactly 1 non-HEAD deployment:
    $deployments_output";
    exit 1; 
}

version_output=$(clasp create-version 'Untitled')
[[ "$version_output" =~ Created\ version\ ([0-9]+) ]] || {
    echo "FAILED. Could not parse version from: $version_output";
    exit 1;
}
version="${BASH_REMATCH[1]}"

clasp update-deployment -V "$version" "$deployment_id"
