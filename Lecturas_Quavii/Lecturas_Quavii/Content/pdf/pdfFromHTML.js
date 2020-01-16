function pdfToHTML(){
  console.log('entro pdf');
var pdf = new jsPDF('landscape');
source = $('#relectura')[0];
specialElementHandlers = {
	'#bypassme': function(element, renderer){
		return true
	}
}
margins = {
    top: 5,
    left: 5,
    width: 2000
  };

pdf.fromHTML(
  	source // HTML string or DOM elem ref.
  	, margins.left // x coord
  	, margins.top // y coord

  	, {
  		'width': 10 // max width of content on PDF
  		, 'elementHandlers': specialElementHandlers
  	}, 
  	function (dispose) {
  	  // dispose: object with X, Y of the last line add to the PDF
  	  //          this allow the insertion of new lines after html
        pdf.save('ResumenLecturas.pdf');
      }
  )		
}