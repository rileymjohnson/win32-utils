import subprocess
import sys

from typing import Set


def get_com_objects() -> Set[str]:
    # I got this powershell script from here:
    # https://powershellmagazine.com/2013/06/27/pstip-get-a-list-of-all-com-objects-available/
    process = subprocess.Popen([
        'powershell.exe',
        'Get-ChildItem',
        'HKLM:\Software\Classes',
        '-ErrorAction SilentlyContinue',
        '| Where-Object {',
        '$_.PSChildName',
        '-match "^\w+\.\w+$"',
        '-and (Test-Path -Path "$($_.PSPath)\CLSID")',
        '} | Select-Object',
        '-ExpandProperty PSChildName'
    ], stdout=subprocess.PIPE)

    output, error = process.communicate()

    if process.returncode != 0 or error is not None:
        raise subprocess.CalledProcessError(
            process.returncode,
            process.args
        )

    return set(output.decode().splitlines())
