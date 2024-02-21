class ArgumentHandler:
    def __init__(self, serialized_arguments):
        """Initialize with the serialized JSON string of arguments."""
        self.defaults = {
            'optional_key': 'default_value',  # Define default values for optional arguments
            # Add more defaults as necessary
        }
        # Deserialize the JSON string to a Python dictionary
        self.arguments = self.deserialize_arguments(serialized_arguments)

    def deserialize_arguments(self, serialized_arguments):
        """Deserialize the JSON string to a Python dictionary."""
        try:
            return json.loads(serialized_arguments)
        except json.JSONDecodeError:
            print("Error decoding JSON arguments.")
            return {}

    def get_argument(self, key):
        """Get an argument value, returning a default if the key is missing."""
        return self.arguments.get(key, self.defaults.get(key))

    def print_all_arguments(self):
        """Print all received options."""
        print(f"All options received: {self.arguments}")

    def print_specific_argument(self, key):
        """Print a specific argument by key, if it exists."""
        if key in self.arguments:
            print(f"{key}: {self.arguments[key]}")
        elif key in self.defaults:
            print(f"{key} (default): {self.defaults[key]}")
        else:
            print(f"{key} not provided and no default value set.")
