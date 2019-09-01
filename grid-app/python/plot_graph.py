"""
Provide a `show` function that can channel a graph to the client.
"""

from pathlib import Path
import os
import base64
import json

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt  # noqa: E402


path_to_temp_file = os.path.join(str(Path.home()), "tmp.svg")


def show():
    plt.savefig(path_to_temp_file)
    with open(path_to_temp_file, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read())

    image_string = str(encoded_string)
    data = {'arguments': ["IMAGE", image_string[2:len(image_string)-1]]}
    data = ''.join(['#IMAGE#', json.dumps(data), '#ENDPARSE#'])

    os.remove(path_to_temp_file)
    print(data, flush=True, end='')
