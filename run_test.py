import matplotlib
matplotlib.use('Agg')

# Patch SHOW_PLOTS to False since no display in test run
import builtins
_orig_open = builtins.open

import sys
sys.path.insert(0, r'd:\New folder')

# Read 1.py, patch TkAgg -> Agg, run it
src = open(r'd:\New folder\1.py').read()
src = src.replace("matplotlib.use('TkAgg')", "matplotlib.use('Agg')")
# Run in namespace so plt.show() renders to file
exec(src, {'__name__': '__main__'})
