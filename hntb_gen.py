import argparse

import glog

from hntb.config_options import read_config
from hntb.facesheets import generate_facesheets
from hntb.ppt import generate_ppt

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Head and neck tumor board file generator.")
    parser.add_argument(
        "--config",
        help="Path to a yaml config file. See `./configs/` directory for example config files.",
        required=True,
    )
    parser.add_argument("--generate", help="Type of file to generate. Options are: [facesheets, ppt].", required=True)
    parser.add_argument("--verbosity", default=0, type=int, help="Verbosity level. Options are: [0, 1].")
    args = parser.parse_args()

    cfg = read_config(args.config)

    if args.verbosity == 0:
        glog.setLevel("INFO")
    elif args.verbosity == 1:
        glog.setLevel("DEBUG")
    else:
        raise NotImplementedError(
            f"\n\nYou passed in '{args.verbosity}' to the --verbosity argument."
            + " It needs to be one of these two: [0, 1].\n"
        )

    if args.generate == "facesheets":
        glog.info("Generating facesheets.")
        generate_facesheets(cfg)
    elif args.generate == "ppt":
        glog.info("Generating powerpoint.")
        generate_ppt(cfg)
    else:
        raise NotImplementedError(
            f"\n\nYou passed in '{args.generate}' to the --generate argument."
            + " It needs to be one of these two: [facesheets, ppt].\n"
        )
