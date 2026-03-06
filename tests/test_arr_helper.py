"""Tests for arr_helper."""

import importlib


def test_package_importable() -> None:
    assert importlib.import_module("arr_helper")
