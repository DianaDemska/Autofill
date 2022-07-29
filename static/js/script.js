function myfunction(id) {
    let firstValue = document.getElementById(id).value;
    secondId = id.replace(/.$/, "O");
    var secondValue = firstValue.replace(/ /g, "-");
    document.getElementById(secondId).value = secondValue;
}

