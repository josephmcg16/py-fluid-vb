"""
py_fluid_vb, provides an interface to access fluid property calculations
from a VB.NET library. It dynamically calls VB.NET functions based on fluid type and property,
allowing for an easy integration between Python and VB.NET to calculate various fluid properties.
"""
import clr
import json
import os


class FluidProperties:
    """
    A class to interface with the FluidProperties VB.NET library, allowing for the calculation
    of various fluid properties.

    Attributes:
        _vb_object (object): An instance of the VB.NET FluidProperties class.
        _vb_funcs_dict (dict): A dictionary mapping fluid types and properties to VB.NET function names.
    """

    def __init__(
        self,
        vb_dll_path=f"{os.path.dirname(os.path.abspath(__file__))}\\bin\\Debug\\net48\\FluidProperties.dll",
    ):
        """
        Initializes the FluidProperties class by loading the VB.NET DLL and reading
        the function lookup from a JSON file.

        Parameters:
            vb_dll_path (str): The file path to the VB.NET DLL. Defaults to the bin\Debug\net48 directory.
        """
        clr.AddReference(vb_dll_path)
        from FluidProperties import FluidProperties

        self._vb_object = FluidProperties()
        with open(f"{os.path.dirname(os.path.abspath(__file__))}\\vb_funcs_lookup.json", "r") as json_file:
            self._vb_funcs_dict = json.load(json_file)

    def _get_vb_func(self, fluid, property):
        """
        Retrieves the VB.NET function based on the fluid and property.

        Parameters:
            fluid (str): The type of fluid for which properties are calculated.
            property (str): The specific property to calculate.

        Returns:
            A callable function from the VB.NET library.
        """
        return getattr(self._vb_object, self._vb_funcs_dict[fluid][property])

    def __call__(self, temperature, pressure, fluid, property):
        """
        Calculates the specified fluid property at the given temperature and pressure.

        Parameters:
            temperature (float): The temperature at which to calculate the property.
            pressure (float): The pressure at which to calculate the property.
            fluid (str): The type of fluid for which properties are calculated.
            property (str): The specific property to calculate.

        Returns:
            The calculated property value as a float.

        Raises:
            ValueError: If the specified fluid or property is not supported.
        """
        if fluid not in self._vb_funcs_dict:
            raise ValueError(
                f"Fluid '{fluid}' not found in the VB functions lookup. Must be one of {list(self._vb_funcs_dict.keys())}"
            )
        if property not in self._vb_funcs_dict[fluid]:
            raise ValueError(
                f"Property '{property}' not found in the VB functions lookup for fluid '{fluid}'. Must be one of {list(self._vb_funcs_dict[fluid].keys())}"
            )

        vb_func = self._get_vb_func(fluid, property)
        return vb_func(temperature, pressure)
