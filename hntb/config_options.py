from dataclasses import dataclass
from pathlib import Path

import yaml


@dataclass
class HNTBConfig:
    # Input files
    active_tumor_board_file: Path

    # Output files
    output_directory: Path
    facesheets_filename: str
    ppt_filename: str

    # Template files
    template_directory: Path
    header_image_filename: str
    facesheet_template_filename: str
    ppt_template_filename: str


def read_config(config_path: str):
    with open(config_path, "r") as f:
        config_dict = yaml.safe_load(f)
    cfg = HNTBConfig(
        active_tumor_board_file=Path(config_dict["active_tumor_board_file"]),
        output_directory=Path(config_dict["output_directory"]),
        facesheets_filename=config_dict["facesheets_filename"],
        ppt_filename=config_dict["ppt_filename"],
        template_directory=Path(config_dict["template_directory"]),
        header_image_filename=config_dict["header_image_filename"],
        facesheet_template_filename=config_dict["facesheet_template_filename"],
        ppt_template_filename=config_dict["ppt_template_filename"],
    )
    return cfg
