"""recombinase — template-guided document synthesis.

Biology: a recombinase is an enzyme that extracts DNA fragments and recombines
them into new molecules using a homologous template as the structural guide.
This package does the same for PowerPoint documents: extract content from
heterogeneous source files, then recombine into a canonical template.
"""

from recombinase.config import TemplateConfig, load_config
from recombinase.generate import generate_deck
from recombinase.inspect import inspect_template

__version__ = "0.2.1"

__all__ = [
    "TemplateConfig",
    "__version__",
    "generate_deck",
    "inspect_template",
    "load_config",
]
