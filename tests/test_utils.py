# Pass the connection string from the CLI
import random
import string
from datetime import datetime, timedelta

import pandas as pd


def get_series_values(series, worksheet):
    """Helper function to extract series values from worksheet."""
    cell_range = series.val.numRef.f
    cells = worksheet[cell_range.split("!")[1]]
    return [cell[0].value for cell in cells]


def get_random_customers():
    def random_date(start, end):
        return start + timedelta(
            seconds=random.randint(0, int((end - start).total_seconds()))
        )

    start_date = datetime(2021, 1, 1)
    end_date = datetime(2023, 12, 31)

    data = {
        "customer_id": list(range(1, 11)),
        "name": [
            "".join(random.choices(string.ascii_uppercase, k=5)) for _ in range(10)
        ],
        "order_date": [random_date(start_date, end_date) for _ in range(10)],
        "payment_amount": [round(random.uniform(10, 1000), 2) for _ in range(10)],
    }

    return pd.DataFrame(data)
