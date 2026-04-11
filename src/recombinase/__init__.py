"""recombinase — template-guided document synthesis.

Biology: a recombinase is an enzyme that extracts DNA fragments and recombines
them into new molecules using a homologous template as the structural guide.
This package does the same for PowerPoint documents: extract content from
heterogeneous source files, then recombine into a canonical template.
"""

from recombinase.config import TemplateConfig, load_config
from recombinase.generate import generate_deck
from recombinase.inspect import inspect_template

__version__ = "0.1.0"

__all__ = [
    "TemplateConfig",
    "load_config",
    "generate_deck",
    "inspect_template",
    "__version__",
]
