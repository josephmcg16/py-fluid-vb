# py-fluid-vb

This Python package provides an interface to calculate fluid properties using a .NET backend. It dynamically calls .NET methods to perform the calculations.

## Getting Started

These instructions will guide you through setting up and installing `py-fluid-vb` on your local machine.

### Prerequisites

- Python 3.x
- .NET SDK (compatible with your .NET project version)
- pip and wheel installed in your Python environment

### Creating the Python Package

Navigate to the Python project directory where setup.py is located.
Create a wheel distribution of the package:

```bash
python setup.py sdist bdist_wheel
```
The wheel file will be created in the dist/ directory.

### Installing the Package

Install the package using the generated wheel file with pip:

```bash
pip install dist/py_fluid_vb-version-py3-none-any.whl
```

### Usage
Import and use the package in your Python scripts:

```python
from myfluidproperties import FluidProperties

# Initialize the class
fluid_props = FluidProperties()

# Call the method
result = fluid_props.get_property(temperature=300, pressure=101325, fluid="nitrogen", property_name="density")
print(result)
```