//Jack Jamieson 2014 NJGS
//This will convert the NJGS Atlas Sheet Coordinates into lat and long in ddmmss.ss and decimal degrees.
//It also finds the USGS quadrangle associated with that ASC.

var errOutput; //Will hold text info about the current error.

//Hold the values created from atlasConversion()
var latDeg;
var lonDeg;
var latDMS;
var lonDMS;

//Hold the quad text;
var quad;


function findQuad(latDD, lonDD)
{
	for(i = 0; i < dat.data.length; i++)
	{
		if(latDD <= dat.data[i][1] && latDD >= dat.data[i][3] && lonDD <= dat.data[i][2] && lonDD >= dat.data[i][4])
		{
			quad = dat.data[i][0];
			return true;
		}
	}
	
	errOutput = "Location outside the defined area.";
	return false;

}

function checkAtlasCoordinate(first, second, third, secondFirstDigit, secondSecondDigit, thirdFirstDigit, thirdSecondDigit, thirdThirdDigit)
{
	//Check for general input errors
    if (first === "" || secondFirstDigit === "" || secondSecondDigit === "" || thirdFirstDigit === "" || thirdSecondDigit === "" || thirdThirdDigit === "") //If the name is empty tell the user
    {

		errOutput = "All fields must be filled in.";
		return false;
    }
	else if(isNaN(first) || isNaN(second) || isNaN(third))//If the user tries to enter something other than a number
	{

		errOutput = "Only numbers allowed.";
		return false;
	}
	else{
	
		//Check for valid block
		if(secondFirstDigit > 4)
		{
	
			errOutput = "Location not on an Atlas Sheet!";
			return false;
		}
		
		if(first == 36)
		{
			if(secondSecondDigit > 6)
			{

				errOutput = "Location not on an Atlas Sheet!";
				return false;

			}
		}
		else if(secondSecondDigit > 5)
		{

			errOutput = "Location not on an Atlas Sheet!";
			return false;

		}
		
		//Check for valid rectangle number
		if(thirdSecondDigit == 0)
		{
			errOutput = "Location not on an Atlas Sheet!";
			return false;
		}
		
		if(thirdThirdDigit == 0)
		{
			errOutput = "Location not on an Atlas Sheet!";
			return false;
		}
	}
	
	return true;
}

function atlasConversion(first, secondFirstDigit, secondSecondDigit, thirdFirstDigit, thirdSecondDigit, thirdThirdDigit)
{
/*
		Latitude and longitude calculation:
		sngSouth = southerning in minutes and sngEast = easting in minutes
		The southerning and easting are determined by successive additions
		of increments as indicated by the given digit of the atlas sheet.
		All increments are measured in minutes of arc. The increment size
		is determined by the given digit. All sngSouth and sngEast increments
		are negative or zero because the reference corner location is the
		northwest corner of the atlas sheet and hence the locations referenced
		to this point are in directions of decreasing latitude and longitude.
		*/
		
		//Calculate southerning and easting of the block coordinate
		var sngSouth = (secondFirstDigit) * (-6);
		var sngEast = (secondSecondDigit - 1) * (-6);
		
		//calculate southerning and easting addition using 1st digit, 3x3 rectangle.
		//increment is 2 minutes
		
		switch(parseInt(thirdFirstDigit)) {
			case 1:
				sngSouth = sngSouth - (0 * 2);
				sngEast = sngEast - (0 * 2);
				break;
			case 2:
				sngSouth = sngSouth - (0 * 2);
				sngEast = sngEast - (1 * 2);
				break;
			case 3:
				sngSouth = sngSouth - (0 * 2);
				sngEast = sngEast - (2 * 2);
				break;
			case 4:
				sngSouth = sngSouth - (1 * 2);
				sngEast = sngEast - (0 * 2);
				break;
			case 5:
				sngSouth = sngSouth - (1 * 2);
				sngEast = sngEast - (1 * 2);
				break;
			case 6:
				sngSouth = sngSouth - (1 * 2);
				sngEast = sngEast - (2 * 2);
				break;
			case 7:
				sngSouth = sngSouth - (2 * 2);
				sngEast = sngEast - (0 * 2);
				break;
			case 8:
				sngSouth = sngSouth - (2 * 2);
				sngEast = sngEast - (1 * 2);
				break;
			case 9:
				sngSouth = sngSouth - (2 * 2);
				sngEast = sngEast - (2 * 2);
				break;
			default:
				errOutput = "Location not on an Atlas Sheet!";
				return false;

		}
		
		//calculate southerning and easting addition using 2nd digit, 3x3 rectangle.
		//increment is 2/3 minutes
		switch(parseInt(thirdSecondDigit)) {
			case 1:
				sngSouth = sngSouth - (0 * 2 / 3);
				sngEast = sngEast - (0 * 2 / 3);
				break;
			case 2:
				sngSouth = sngSouth - (0 * 2 / 3);
				sngEast = sngEast - (1 * 2 / 3);
				break;
			case 3:
				sngSouth = sngSouth - (0 * 2 / 3);
				sngEast = sngEast - (2 * 2 / 3);
				break;
			case 4:
				sngSouth = sngSouth - (1 * 2 / 3);
				sngEast = sngEast - (0 * 2 / 3);
				break;
			case 5:
				sngSouth = sngSouth - (1 * 2 / 3);
				sngEast = sngEast - (1 * 2 / 3);
				break;
			case 6:
				sngSouth = sngSouth - (1 * 2 / 3);
				sngEast = sngEast - (2 * 2 / 3);
				break;
			case 7:
				sngSouth = sngSouth - (2 * 2 / 3);
				sngEast = sngEast - (0 * 2 / 3);
				break;
			case 8:
				sngSouth = sngSouth - (2 * 2 / 3);
				sngEast = sngEast - (1 * 2 / 3);
				break;
			case 9:
				sngSouth = sngSouth - (2 * 2 / 3);
				sngEast = sngEast - (2 * 2 / 3);
				break;
			default:
				errOutput = "Location not on an Atlas Sheet!";
				return false;
		}
		
		//calculate southerning and easting addition using 3rd digit, 3x3 rectangle.
		//increment is 2/9 minutes
		switch (parseInt(thirdThirdDigit)) {
			case 1:
				sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (0 * 2 / 9) - (1 / 9);
				break;
			case 2:
				sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (1 * 2 / 9) - (1 / 9);
				break;
			case 3:
				sngSouth = sngSouth - (0 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (2 * 2 / 9) - (1 / 9);
				break;
			case 4:
				sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (0 * 2 / 9) - (1 / 9);
				break;
			case 5:
				sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (1 * 2 / 9) - (1 / 9);
				break;
			case 6:
				sngSouth = sngSouth - (1 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (2 * 2 / 9) - (1 / 9);
				break;
			case 7:
				sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (0 * 2 / 9) - (1 / 9);
				break;
			case 8:
				sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (1 * 2 / 9) - (1 / 9);
				break;
			case 9:
				sngSouth = sngSouth - (2 * 2 / 9) - (1 / 9);
				sngEast = sngEast - (2 * 2 / 9) - (1 / 9);
				break;
			default:
				errOutput = "Location not on an Atlas Sheet!";
				return false;
		}
		
			
		//Add easting and southerning to the northwest corner of the appropriate
		//atlas sheet. The northwest corner latitude and longitude is in minutes.
		switch(parseInt(first)) {
			case 21:
				sngSouth = sngSouth + 2484;
				sngEast = sngEast + 4512;
				break;
			case 22:
				sngSouth = sngSouth + 2484;
				sngEast = sngEast + 4486;
				break;
			case 23:
				sngSouth = sngSouth + 2484;
				sngEast = sngEast + 4460;
				break;
			case 24:
				sngSouth = sngSouth + 2456;
				sngEast = sngEast + 4512;
				break;
			case 25:
				sngSouth = sngSouth + 2456;
				sngEast = sngEast + 4486;
				break;
			case 26:
				sngSouth = sngSouth + 2456;
				sngEast = sngEast + 4460;
				break;
			case 27:
				sngSouth = sngSouth + 2428;
				sngEast = sngEast + 4512;
				break;
			case 28:
				sngSouth = sngSouth + 2428;
				sngEast = sngEast + 4486;
				break;
			case 29:
				sngSouth = sngSouth + 2428;
				sngEast = sngEast + 4460;
				break;
			case 30:
				sngSouth = sngSouth + 2400;
				sngEast = sngEast + 4538;
				break;
			case 31:
				sngSouth = sngSouth + 2400;
				sngEast = sngEast + 4512;
				break;
			case 32:
				sngSouth = sngSouth + 2400;
				sngEast = sngEast + 4486;
				break;
			case 33:
				sngSouth = sngSouth + 2400;
				sngEast = sngEast + 4460;
				break;
			case 34:
				sngSouth = sngSouth + 2372;
				sngEast = sngEast + 4538;
				break;
			case 35:
				sngSouth = sngSouth + 2372;
				sngEast = sngEast + 4512;
				break;
			case 36:
				sngSouth = sngSouth + 2372;
				sngEast = sngEast + 4486;
				break;
			case 37:
				sngSouth = sngSouth + 2344;
				sngEast = sngEast + 4500;
				break;
			default:
				errOutput = "Location not on an Atlas Sheet!";
				return false;
				
		}
		
		
		//Convert to decimal degrees
		var decLat = Number(sngSouth / 60).toFixed(3);
		var decLong = Number(sngEast / 60).toFixed(3);
		
		latDeg = decLat;
		lonDeg = decLong;
		
		/*
		if(causedErr == false)
			divDeg.innerHTML = "<strong>Decimal Degrees (NAD27): </strong>" + "Latitude: " + decLat + " " + "Longitude: -" + decLong;
		else 
			divDeg.innerHTML = "<strong>Decimal Degrees (NAD27): </strong>";
		*/
		
		//Convert to ddmmss.ss
		var degLatD = ~~(decLat);

		var degLatM = ~~((decLat - (~~(decLat))) * 60);//bitwise calc to get the whole number
		var degLatS = (((decLat - (~~(decLat))) * 60) - degLatM) * 60;//bitwise calc to get the whole number
		
		var degLonD = ~~(decLong);
		var degLonM = ~~((decLong - (~~(decLong))) * 60);//bitwise calc to get the whole number
		var degLonS = (((decLong - (~~(decLong))) * 60) - degLonM) * 60;//bitwise calc to get the whole number
		
		degLatS = Number(degLatS).toFixed(2);
		degLonS = Number(degLonS).toFixed(2);

		//Fix numbers under 10 without leading 0
		if(degLatM < 10)
		{
			degLatM = "" + degLatM;
			degLatM = "0" + degLatM;
		}
		if(degLatS < 10)
		{
			degLatS = "" + degLatS;
			degLatS = "0" + degLatS;
		}
		if(degLonM < 10)
		{
			degLonM = "" + degLonM;
			degLonM = "0" + degLonM;
		}
		if(degLonS < 10)
		{
			degLonS = "" + degLonS;
			degLonS = "0" + degLonS;
		}	
	
		latDMS = degLatD + "" + degLatM + "" + degLatS + "";
		lonDMS = degLonD + "" + degLonM + "" + degLonS + "";
		/*
		if(causedErr == false)
			divDMS.innerHTML = "<strong>ddmmss.ss (NAD27):</strong>" + "Latitude: " + degLatD + "" + degLatM + "" + degLatS + " Longitude: -" + degLonD + "" + degLonM + "" + degLonS + "";
		else 
			divDMS.innerHTML = "<strong>ddmmss.ss (NAD27):</strong>";
	
		*/
		/*
		var quad = (findQuad(Number("" + degLatD + degLatM + degLatS), Number("" + degLonD + degLonM + degLonS)));
		//console.log(findQuad(Number("" + degLatD + degLatM + degLatS), Number("" + degLonD + degLonM + degLonS)));
		//var quad = (findQuad(400353, 743953));

		if(causedErr == false)
			divQuad.innerHTML = "<strong>USGS Quadrangle: </strong>" + quad;
		*/
		
		return true;
} 
		
