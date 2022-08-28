import yaml
import pandas as pd
import re

class FillNPrint:

    def parse_yaml(self, file): #parse yaml files
        with open(file, 'r') as stream:
            try:
                return yaml.safe_load(stream)
            except yaml.YAMLError as exc:
                return "error: invalid yaml file"

        return df

