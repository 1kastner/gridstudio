"""
Here, the following files are pulled together:
- datasheet
- plot_graph

The external interface is presented:
- sheet: Read and write to a range in the excel sheet
- cell: Read and write to a cell in the excel sheet
- plot: Visually display whatever has been previously plotted with matplotlib
"""

import traceback
import warnings

from plot_graph import show  # noqa: F401
from datasheet import sheet, cell  # noqa: F401


warnings.filterwarnings("ignore", message="numpy.dtype size changed")
warnings.filterwarnings("ignore", message="numpy.ufunc size changed")


def parseCall(*arg):
    result = ""
    try:
        if len(arg) > 1:
            eval_result = eval(arg[0] + "(\""+'","'.join(arg[1:])+"\")")
        else:
            eval_result = eval(arg[0] + "()")

        if isinstance(eval_result, numbers.Number) \
                and not isinstance(eval_result, bool):
            result = str(eval_result)
        else:
            result = "\"" + str(eval_result) + "\""

    except (RuntimeError, TypeError, NameError) as e:
        result = "\"" + "Unexpected error:" + str(e.message) + "\""

    print("#PYTHONFUNCTION#"+result+"#ENDPARSE#", flush=True, end='')


def getAndExecuteInput():
    command_buffer = ""
    while True:
        code_input = input("")
        # when empty line is found, execute code
        if code_input == "":
            try:
                exec(command_buffer, globals(), globals())

                # only print COMANDCOMPLETE when the input doesn't
                # start with parseCall, since it's a special internal Python
                # call which requires a single print between exec and return
                if not command_buffer.startswith("parseCall"):
                    print("#COMMANDCOMPLETE##ENDPARSE#", end='', flush=True)
            except:
                traceback.print_exc()
            command_buffer = ""
        else:
            command_buffer += code_input + "\n"


# testing
# sheet("A1:A2", [1,2])
# df = pd.DataFrame({'a':[1,2,3], 'b':[4,5,6]})
# sheet("A1:B2")

getAndExecuteInput()
