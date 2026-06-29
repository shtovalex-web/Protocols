# -*- coding: utf-8
"""Обратная совместимость: update_success → desktop_updater."""

from tkinter import messagebox

import protocol_updater_config  # noqa: F401

from desktop_updater.installer import exit_after_update  # noqa: F401
from desktop_updater.success import notify_update_success_and_exit  # noqa: F401
