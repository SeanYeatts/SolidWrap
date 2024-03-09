'''Copyright (c) 2024 Sean Yeatts. All rights reserved.'''

from __future__ import annotations


# IMPORTS - STANDARD LIBRARY
import logging


# SYMBOLS
__all__ = [
    'log',
]


# Setup module logger
log = logging.getLogger(__name__)
log.setLevel(logging.INFO)

# Configure console handler
handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)

# Configure formatter
string = f'%(asctime)s %(levelname)s - %(message)s'
date = f'%H:%M:%S'
formatter = logging.Formatter(string, datefmt=date)

# Assign configured parameters
handler.setFormatter(formatter)
log.addHandler(handler)
