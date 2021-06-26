// Calculate distance betwwen 2 points on earth

const latitude1 = 53.32055555555556;
const longitude1 = -1.7297222222222221;

const latitude2 = 53.31861111111111;
const longitude2 = -1.6997222222222223;

function toRadians(degree)
{
    return (Math.PI / 180) * degree;
}

function distance(lat1, long1, lat2, long2)
{
    // Convert the latitudes
    // and longitudes
    // from degree to radians.

    lat1 = toRadians(lat1);
    long1 = toRadians(long1);
    lat2 = toRadians(lat2);
    long2 = toRadians(long2);

    let dlong = long2 - long1;
    let dlat = lat2 - lat1;

    let ans = Math.pow(Math.sin(dlat / 2), 2) +
              Math.cos(lat1) * Math.cos(lat2) *
              Math.pow(Math.sin(dlong / 2), 2);

    ans = 2 * Math.asin(Math.sqrt(ans));

    // Radius of Earth in Kilometers, R = 6371
    // Use R = 3956 for miles
    const R = 6371;

    // Calculate the result
    ans = ans * R;

    return ans;
}

let d = distance(latitude1, longitude1,
                 latitude2, longitude2);

console.log("Distance:", d, "km");
