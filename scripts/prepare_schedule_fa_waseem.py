#!/usr/bin/env python3
"""Run both Mohammad Waseem Individual + Joint Schedule FA generators."""

from prepare_schedule_fa_waseem_individual import run as run_individual
from prepare_schedule_fa_waseem_joint import run as run_joint

if __name__ == "__main__":
    run_individual()
    run_joint()
