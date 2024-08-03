
  /**
* Get Distance between 2 different addresses.
* @param start_address Address as string Ex. "300 N LaSalles St, Chicago, IL"
* @param end_address Address as string Ex. "900 N LaSalles St, Chicago, IL"
* @param return_type Return type as string Ex. "miles" or "kilometers" or "minutes" or "hours"
* @customfunction
*/

function GOOGLEMAPS(start_address,end_address,return_type) {

  // https://www.chicagocomputerclasses.com/
  // Nov 2017
  // improvements needed
  
  var mapObj = Maps.newDirectionFinder();
  mapObj.setOrigin(start_address);
  mapObj.setDestination(end_address);
  var directions = mapObj.getDirections();
  
  var getTheLeg = directions["routes"][0]["legs"][0];
  
  var meters = getTheLeg["distance"]["value"];
  
  switch(return_type){
    case "miles":
      return meters * 0.000621371;
      break;
    case "minutes":
        // get duration in seconds
        var duration = getTheLeg["duration"]["value"];
        //convert to minutes and return
        return duration / 60;
      break;
    case "hours":
        // get duration in seconds
        var duration = getTheLeg["duration"]["value"];
        //convert to hours and return
        return duration / 60 / 60;
      break;      
    case "kilometers":
      return meters / 1000;
      break;
    default:
      return "Error: Wrong Unit Type";
   }
  
}

