//datetime = JobRole
//city = EducationField
//state = BusinessTravel
//cntry = YearsWithCurrManager
//shape = JobInvolvement
//


// Get references to the tbody element, input field and button
var $tbody = document.querySelector('tbody');
var $JobRole = document.querySelector('#date');
var $EducationField = document.querySelector('#EducationField');
var $BusinessTravel = document.querySelector('#BusinessTravel');
var $cntry = document.querySelector('#country');
var $shape = document.querySelector('#shape');
var $rpp = document.querySelector('#page');


// Add an event listener to the searchButton, call handleSearchButtonClick when clicked
$JobRole.addEventListener('change', handleSearchButtonClick);
$EducationField.addEventListener('change', handleSearchButtonClick);
$BusinessTravel.addEventListener('change', handleSearchButtonClick);
$cntry.addEventListener('change', handleSearchButtonClick);
$shape.addEventListener('change', handleSearchButtonClick);
$rpp.addEventListener('change', getPagination);

// Set filteredData to dataSet initially
// dataSet comes from the data.js file
// var filteredData = dataSet.slice(0,1000);
var filteredData = dataSet;

// renderSearch renders the filter crieteria for Country, Shape and Results per Page
var countryList = [];
var shapeList = [];
var prevPage = 0;
var nextPage = 0;
var page_count = 0;

function renderSearch()
{
  for (var i = 0; i < filteredData.length; i++)
  {
    countryList[i] = filteredData[i].country;
    shapeList[i] = filteredData[i].shape;
  }

  var uniqueCountryList = Array.from(new Set(countryList)).sort()
  var uniqueShapeList = Array.from(new Set(shapeList)).sort()

  // populate Country dropdown from dataset
  d3.select("#country").append("option").text('all');
  for (i=0; i < uniqueCountryList.length; i++)
    d3.select("#country").append("option").text(uniqueCountryList[i]).attr("value",uniqueCountryList[i]);

  // populate Shape dropdown from dataset
  d3.select("#shape").append("option").text('all');
  for (i=0; i < uniqueShapeList.length; i++)
    d3.select("#shape").append("option").text(uniqueShapeList[i]).attr("value",uniqueShapeList[i]);

  for (i=10; i >= 1; i--)
    d3.select("#page").append("option").text(i*10).attr("value",i*10);
}

renderSearch();

// add pagination links
function getPagination()
{
  var row_count = $rpp.value.trim();
  if (!row_count)
    row_count = defaultRowCount;

  console.log("row_count" + row_count);
  console.log("Total:"+ filteredData.length);

  var $ul = d3.select('.pagination');
  d3.select('.pagination').selectAll("li").remove();

  if (row_count > 0 && filteredData.length > 0 )
  {
      page_count = Math.ceil(filteredData.length / row_count);
      console.log("pages:" + page_count);
      $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text('Previous');

      if (page_count <= 15)
      {
        for (var i=1; i<=page_count; i++)
          $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text(i);
      }
      else
      {
        for (i=1; i<=10; i++)
          $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text(i);

        $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text('.....');

        for (i=page_count-10; i<=page_count; i++)
          $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text(i);
      }
      $ul.append('li').attr('class', 'page-item').append('a').attr('class','page-link').text('Next');

  }
  addEvent();
  displayData(1, row_count);
}

getPagination();

// add Event Handlers for each <li>
function addEvent()
{
  var links = document.getElementsByTagName('a');

    for (var i = 0; i < links.length; i++) {
      var link = links[i];
      link.addEventListener('click', myFunction);
    }
}

// Event Listener callback function to display data for the selected page
function myFunction() {
  var row_count = $rpp.value.trim();
  if (!row_count)
    row_count = defaultRowCount;
  if (this.text == 'Previous')
  {
    if (prevPage > 1)
      displayData(prevPage--, row_count)
    else if (prevPage == 1)
      displayData(1, row_count)
  }
  else if(this.text == 'Next')
  {
    if (nextPage < page_count)
      displayData(nextPage++, row_count)
    else if (nextPage == page_count)
      displayData(page_count, row_count)
  }
  else
  {
    prevPage = this.text;
    nextPage = this.text;
    displayData(this.text, row_count)
  }
}

// display data for the selected page
function displayData(p, rows) {
  console.log("page no.:"+ p);
  console.log("row count:"+ rows);
  $tbody.innerHTML = '';

  if (parseInt(Number(p)) == p)
  {
      data = filteredData.slice((p-1)*rows, (p*rows));
      console.log(data);
  }
  else
  {
	   alert('Please click on Previous or Next to navigate the pages.');
     console.log('Sorry!! This is not hanlded :( Need some more time.');
     return;
  }

  for (var i = 0; i < data.length; i++) {
    var UFOData = data[i];
    var fields = Object.keys(UFOData);
    // Create a new row in the tbody, set the index to be i + startingIndex
    var $row = $tbody.insertRow(i);
    for (var j = 0; j < fields.length; j++) {
      // For every field in the UFOData object, create a new cell at set its inner text to be the current value
      // at the current UFOData's field
      var field = fields[j];
      var $cell = $row.insertCell(j);
      $cell.innerText = UFOData[field];
    }
  }
}

// renderTable renders the filteredData to the tbody
function renderTable() {
  $tbody.innerHTML = '';
  for (var i = 0; i < filteredData.length; i++) {
    // Get get the current UFOData object and its fields
    var UFOData = filteredData[i];
    var fields = Object.keys(UFOData);
    // Create a new row in the tbody, set the index to be i + startingIndex
    var $row = $tbody.insertRow(i);
    for (var j = 0; j < fields.length; j++) {
      // For every field in the UFOData object, create a new cell at set its inner text to be the current value
      // at the current UFOData's field
      var field = fields[j];
      var $cell = $row.insertCell(j);
      $cell.innerText = UFOData[field];
    }
  }
}

function filterByDate(n, Date)
{
  if (n == 0)
  {
   // Set filteredData to an array of all UFOData whose "Date" matches the filter
   filteredData = dataSet.filter(function(UFOData) {
    var dataSetData = UFOData.JobRole.toLowerCase();
   // If true, add the UFOData to the filteredData, otherwise don't add it to filteredData
      return dataSetData === Date;   });
  }
  else
  {
   // Set filteredData to an array of all UFOData whose "Date" matches the filter
   filteredData = filteredData.filter(function(UFOData) {
    var dataSetData = UFOData.JobRole.toLowerCase();
      // If true, add the UFOData to the filteredData, otherwise don't add it to filteredData
      return dataSetData === Date;    });
  }
}

function filterByEducationField(n, EducationField)
{
  if (n == 0)
  {
    filteredData = dataSet.filter(function(UFOData) {
      var dataSetData= UFOData.EducationField.toLowerCase();
      return dataSetData === EducationField; });
  }
  else
  {
    filteredData = filteredData.filter(function(UFOData) {
      var dataSetData= UFOData.EducationField.toLowerCase();
      return dataSetData === EducationField; });
  }
}

function filterByBusinessTravel(n, BusinessTravel)
{
  if (n == 0)
  {
    filteredData = dataSet.filter(function(UFOData) {
      var dataSetData= UFOData.BusinessTravel.toLowerCase();
      return dataSetData === BusinessTravel; });
  }
  else
  {
    filteredData = filteredData.filter(function(UFOData) {
      var dataSetData= UFOData.BusinessTravel.toLowerCase();
      return dataSetData === BusinessTravel; });
  }
}

function filterByCountry(n, Country)
{
  if (n == 0)
  {
      filteredData = dataSet.filter(function(UFOData) {
        var dataSetData= UFOData.country.toLowerCase();
        return dataSetData === Country; });
    }
    else
    {
      filteredData = filteredData.filter(function(UFOData) {
        var dataSetData= UFOData.country.toLowerCase();
        return dataSetData === Country; });
    }
  }

function filterByShape(n, Shape)
{
    if (n == 0)
    {
      filteredData = dataSet.filter(function(UFOData) {
        var dataSetData= UFOData.shape.toLowerCase();
        return dataSetData === Shape; });
    }
    else
    {
      filteredData = filteredData.filter(function(UFOData) {
        var dataSetData= UFOData.shape.toLowerCase();
        return dataSetData === Shape; });
    }
  }

function handleSearchButtonClick() {
  // Format the user's search by removing leading and trailing whitespace, lowercase the string
  var filterJobRole= $JobRole.value.trim().toLowerCase();
  console.log("filterJobRole:"+filterJobRole)
  var filterEducationField= $EducationField.value.trim().toLowerCase();
  console.log("filterEducationField:"+filterEducationField)
  var filterBusinessTravel = $BusinessTravel.value.trim().toLowerCase();
  console.log("filterBusinessTravel:" + $BusinessTravel.value)
  var filterCountry = $cntry.value.trim().toLowerCase();
  console.log("filterCountry:" + $cntry.value)
  var filterShape = $shape.value.trim().toLowerCase();
  console.log("filterShape:" + $shape.value)

  if (filterCountry == 'all' && filterShape == 'all')
  {
        if (filterJobRole != '' && filterEducationField != '' && filterBusinessTravel != '' )
        {
          filterByDate(0, filterJobRole);
          filterByEducationField(1, filterEducationField);
          filterByBusinessTravel(1, filterBusinessTravel);
        }
        else if(filterJobRole == '' && filterEducationField != '' && filterBusinessTravel != '')
        {
          filterByEducationField(0, filterEducationField);
          filterByBusinessTravel(1, filterBusinessTravel);
        }
        else if(filterJobRole == '' && filterEducationField == '' && filterBusinessTravel != '')
        {
          filterByBusinessTravel(0, filterBusinessTravel);
        }
  }
  else if(filterCountry != 'all' && filterShape == 'all')
    filterByCountry(0, filterCountry);
  else if(filterCountry == 'all' && filterShape != 'all')
    filterByShape(0, filterShape);
  else if(filterCountry != 'all' && filterShape != 'all')
  {
    filterByCountry(0, filterCountry);
    filterByShape(1, filterShape);
  }

  if (filterJobRole != '')
    filterByDate(1, filterJobRole);

  if (filterEducationField != '')
    filterByEducationField(1, filterEducationField);

  if (filterBusinessTravel != '')
    filterByBusinessTravel(1, filterBusinessTravel);

  if (filterCountry != '' && filterCountry != 'all')
    filterByCountry(1, filterCountry);

  if (filterShape != '' && filterShape != 'all')
    filterByShape(1, filterShape);

    getPagination();
}

// Render the table for the first time on page load
// renderTable();
