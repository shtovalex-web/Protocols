# -*- coding: utf-8
"""Обратная совместимость: update_manifest → desktop_updater."""

import protocol_updater_config  # noqa: F401

from desktop_updater.manifest import (  # noqa: F401
    DataFilePayload,
    UpdateManifest,
    UpdateManifestError,
    WindowsUpdatePayload,
    load_update_manifest,
    sha256_file,
)
