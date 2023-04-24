import requests
import pandas as pd
from geopy.geocoders import Nominatim

# Replace YOUR_ACCESS_TOKEN with your Strava access token
access_token = '9739f7b3b1c683865f92e6c3d04767b9fdefdcc3'

# Make a request to get the user's activities
url = 'https://www.strava.com/api/v3/athlete/activities?per_page=50&page=2'
headers = {'Authorization': f'Bearer {access_token}'}
response = requests.get(url, headers=headers)
activities = response.json()

# http://www.strava.com/oauth/authorize?client_id=104674&response_type=code&redirect_uri=http://localhost/exchange_token&approval_prompt=force&scope=read_all
# 670417e0491265171fffb17a359e14603f1755d4
# http://www.strava.com/oauth/authorize?client_id=104674&response_type=code&redirect_uri=http://localhost/exchange_token&approval_prompt=force&scope=activity:read_all
# 670417e0491265171fffb17a359e14603f1755d4
# http://www.strava.com/oauth/token?client_id=104674&client_secret=4d5466440f2ae00e83e2827b34c3927f24123fba&code=670417e0491265171fffb17a359e14603f1755d4&grant_type=authorization_code

# Create a dictionary to store the total kilometers for each country
km_by_country = {}

# Create a geolocator using Nominatim
geolocator = Nominatim(user_agent='strava')

# Loop through the activities and add up the kilometers for each country
for activity in activities:
    # Make a request to get the activity details
    url = f'https://www.strava.com/api/v3/activities/{activity}'
    # activity_Id = activity.name
    # url = f'https://www.strava.com/api/v3/activities/activity.name'

    response = requests.get(url, headers=headers)
    activity_details = response.json()
    print(activity_details)
    # Extract the country and distance for the activity
    # country = activity_details['location_country']
    # country = activity_details['timezone']
    # distance = activity_details['distance']
    
    # Use reverse geocoding to get the country for the latitude and longitude

     # Extract the latitude and longitude for the activity
    if len(activity_details['start_latlng']):
        lat_lng = activity_details['start_latlng']
        location = geolocator.reverse(lat_lng)
        # print(location.raw['address']['state'])
        country = location.raw['address']['country_code']
                # Add the distance to the country's total
        distance = activity_details['distance']
        # Add the distance to the country's total
        if country in km_by_country:
            km_by_country[country] += distance /1000
        else:
            km_by_country[country] = distance /1000

# Convert the dictionary to a pandas DataFrame and sort by total kilometers
df = pd.DataFrame.from_dict(km_by_country, orient='index', columns=['total_km'])
df = df.sort_values(by='total_km', ascending=False)

# Print the results
print(df)
