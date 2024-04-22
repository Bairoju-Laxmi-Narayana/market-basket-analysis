document.getElementById("aprioriBtn").addEventListener("click", function() {
    executePythonProgram("C:/Users/91934/Documents/major project/apriori.py");
});

document.getElementById("fpGrowthBtn").addEventListener("click", function() {
    executePythonProgram("C:/Users/91934/Documents/major project/fpgrowth.py");
});

document.getElementById("eclatBtn").addEventListener("click", function() {
    executePythonProgram("C:/Users/91934/Documents/major project/eclat.py");
});

function executePythonProgram(path) {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (this.readyState == 4 && this.status == 200) {
            document.getElementById("result").innerHTML = this.responseText;
        }
    };
    xhttp.open("GET", "execute_program.py?path=" + path, true);
    xhttp.send();
}
