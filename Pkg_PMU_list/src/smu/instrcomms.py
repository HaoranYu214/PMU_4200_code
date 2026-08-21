"""Minimal PyVISA communication wrapper for the 4200A-SCS."""

from __future__ import annotations

import pyvisa as visa


class Communications:
    """Open, query, and close one VISA instrument resource."""

    def __init__(self, instrument_resource_string=None):
        self._instrument_resource_string = instrument_resource_string
        self._resource_manager = visa.ResourceManager()
        self._instrument_object = None
        self._timeout = 20000
        self._echo_cmds = False

    def connect(self, instrument_resource_string=None, timeout=None):
        if instrument_resource_string is not None:
            self._instrument_resource_string = instrument_resource_string
        self._instrument_object = self._resource_manager.open_resource(
            self._instrument_resource_string
        )
        self._instrument_object.timeout = self._timeout if timeout is None else timeout
        if timeout is not None:
            self._timeout = timeout
        return self

    def disconnect(self):
        if self._instrument_object is not None:
            self._instrument_object.close()
            self._instrument_object = None

    def close(self):
        """Close the instrument and VISA resource manager."""
        self.disconnect()
        if self._resource_manager is not None:
            self._resource_manager.close()
            self._resource_manager = None

    def query(self, command):
        if self._instrument_object is None:
            raise RuntimeError("No instrument connection is open.")
        if self._echo_cmds:
            print(command)
        return self._instrument_object.query(command).rstrip()
