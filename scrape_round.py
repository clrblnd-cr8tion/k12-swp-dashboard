#!/usr/bin/env python3
"""
Helper script to generate the scraping commands for a round.
This outputs the plan IDs and URLs that need to be visited.
The actual browser interaction happens via the Chrome extension in the conversation.
"""
import json
import sys

# Round configurations
ROUNDS = {
    "R5": {
        "round": 5,
        "yearCode": "2023004",
        "grantFY": "2022-23",
        "quarterlyFYs": ["2023004", "2024004", "2025004"],
        "quarterlyLabels": ["FY22-23_Q4", "FY23-24_Q2", "FY23-24_Q4", "FY24-25_Q2", "FY24-25_Q4"],
        "quarterlyFYMap": {
            "FY22-23_Q4": "2023004",
            "FY23-24_Q2": "2024004",
            "FY23-24_Q4": "2024004",
            "FY24-25_Q2": "2025004",
            "FY24-25_Q4": "2025004"
        },
        "planIds": ['20660', '20672', '20699', '20683', '20697', '20616', '20692', '20593', '20615', '20606', '20618', '20626', '20654', '20611', '20627', '20619', '20592', '20603', '20635', '20590', '20690', '20602', '20591', '20624', '20625', '20655', '20588', '20589']
    },
    "R6": {
        "round": 6,
        "yearCode": "2024004",
        "grantFY": "2023-24",
        "quarterlyFYs": ["2024004", "2025004", "2026004"],
        "quarterlyLabels": ["FY23-24_Q4", "FY24-25_Q2", "FY24-25_Q4", "FY25-26_Q2", "FY25-26_Q4"],
        "quarterlyFYMap": {
            "FY23-24_Q4": "2024004",
            "FY24-25_Q2": "2025004",
            "FY24-25_Q4": "2025004",
            "FY25-26_Q2": "2026004",
            "FY25-26_Q4": "2026004"
        },
        "planIds": ['25883', '25896', '25679', '25856', '25636', '25760', '25743', '25669', '25858', '25686', '25751', '25637', '25638', '25843', '25903', '25685', '25643', '25641', '25651', '25668', '25854', '25646', '25882', '25678', '25658', '25656', '25649', '25632', '25633', '25757']
    },
    "R7": {
        "round": 7,
        "yearCode": "2025004",
        "grantFY": "2024-25",
        "quarterlyFYs": ["2025004", "2026004", "2027004"],
        "quarterlyLabels": ["FY24-25_Q4", "FY25-26_Q2", "FY25-26_Q4", "FY26-27_Q2", "FY26-27_Q4"],
        "quarterlyFYMap": {
            "FY24-25_Q4": "2025004",
            "FY25-26_Q2": "2026004",
            "FY25-26_Q4": "2026004",
            "FY26-27_Q2": "2027004",
            "FY26-27_Q4": "2027004"
        },
        "planIds": ['29787', '29752', '29789', '29754', '29762', '29843', '29803', '29763', '29775', '29776', '29788', '29760', '29799', '29800', '29796', '29765', '29755', '29757', '29802', '29816', '29756', '29769', '29770', '29790', '29767', '29764', '29753', '29878', '29810', '29751', '29842', '29841', '29773', '29794', '29780', '29781', '29761', '29792']
    },
    "R8": {
        "round": 8,
        "yearCode": "2026004",
        "grantFY": "2025-26",
        "quarterlyFYs": ["2026004", "2027004", "2028004"],
        "quarterlyLabels": ["FY25-26_Q4", "FY26-27_Q2", "FY26-27_Q4", "FY27-28_Q2", "FY27-28_Q4"],
        "quarterlyFYMap": {
            "FY25-26_Q4": "2026004",
            "FY26-27_Q2": "2027004",
            "FY26-27_Q4": "2027004",
            "FY27-28_Q2": "2028004",
            "FY27-28_Q4": "2028004"
        },
        "planIds": ['32642', '32740', '32741', '32876']
    }
}

if __name__ == "__main__":
    round_name = sys.argv[1] if len(sys.argv) > 1 else "R7"
    config = ROUNDS[round_name]

    print(f"Round: {round_name}")
    print(f"Grant count: {len(config['planIds'])}")
    print(f"Year code: {config['yearCode']}")
    print(f"Quarterly FYs: {config['quarterlyFYs']}")
    print(f"\nURLs to visit per grant: {1 + len(config['quarterlyFYs'])} (1 primary + {len(config['quarterlyFYs'])} quarterly)")
    print(f"Total page loads: {len(config['planIds']) * (1 + len(config['quarterlyFYs']))}")

    # Print first 3 as examples
    for pid in config['planIds'][:3]:
        print(f"\nGrant {pid}:")
        print(f"  Primary: /swpk/fiscal-reports/plans/{pid}?duration={config['yearCode']}")
        for fy in config['quarterlyFYs']:
            print(f"  Quarterly: /swpk/fiscal-reports/plans/{pid}?duration={fy}")
