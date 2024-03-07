from datetime import datetime, timedelta

def date_reader(date):
    # Extract milliseconds and covert to sec
    timestamp_ms = int(date.strip('/Date()'))
    timestamp_s = timestamp_ms / 1000.0
    # Convert to date
    finalDate = datetime.utcfromtimestamp(timestamp_s)

    return finalDate

def closest_date(date_list):
    now = datetime.utcnow()
    closest_date = min(date_list, key=lambda date: abs(date-now))
    return closest_date.strftime('%Y-%m-%d %H:%M:%S')

def read_closest_date(datasets):
    '''Returns the closest date for a list of ressources'''
    allCreatedDate = []
    for dataset in datasets.values():
        allCreatedDate.append(date_reader(dataset["CreatedDate"]))
    return closest_date(allCreatedDate)
    