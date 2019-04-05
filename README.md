Cluster Export
==============

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](http://www.apache.org/licenses/LICENSE-2.0) ![This script was last tested in Nuix 7.8](https://img.shields.io/badge/Script%20Tested%20in%20Nuix-7.8-green.svg)

View the GitHub project [here](https://github.com/Nuix/Cluster-Export) or download the latest release [here](https://github.com/Nuix/Cluster-Export/releases).

# Overview

This script exports items into separate folders for each cluster selected from a cluster run.

# Getting Started

## Setup

Begin by downloading the latest release of this code.  Extract the contents of the archive into your Nuix scripts directory.  In Windows the script directory is likely going to be either of the following:

- `%appdata%\Nuix\Scripts` - User level script directory
- `%programdata%\Nuix\Scripts` - System level script directory

# Usage

Before running the script you will need to have an open Nuix case containing clusters.

Running the script will display a ClusterExportSettings dialog, prompting the user to select an export directory and a cluster run and its clusters. Each selected cluster will have its items exported to `..\<Cluster Run>\<Cluster ID>\<Item Name>.<Extension>`.

Cluster IDs for pseudoclusters will be replaced with Strings. This means:
- Cluster ID -1 is 'unclusterable'
- Cluster ID -2 is 'ignorable'

The item's extension will be the first non-blank value from:
- Item.getOriginalExtension
- Item.getCorrectedExtension
- Item.getType.getPreferredExtension
- 'BIN'

# Cloning this Repository

This script relies on code from [Nx](https://github.com/Nuix/Nx) to present a settings dialog and progress dialog.  This JAR file is not included in the repository (although it is included in release downloads).  If you clone this repository, you will also want to obtain a copy of Nx.jar by either:
1. Building it from [the source](https://github.com/Nuix/Nx)
2. Downloading an already built JAR file from the [Nx releases](https://github.com/Nuix/Nx/releases)

Once you have a copy of Nx.jar, make sure to include it in the same directory as the script.

# License

```
Copyright 2019 Nuix

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
```
