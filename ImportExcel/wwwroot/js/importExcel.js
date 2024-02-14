
let excelInterval;
document.getElementById("submitExcel").style.display = "none";
function SubmitExcel() {
    debugger;
    clearInterval(excelInterval);
    let excel = document.getElementById("Excel").files[0];
    let formData = new FormData();
    formData.append("excel", excel);
    fetch('/Home/AddCustomersByExcel', { method: "POST", body: formData });
    document.getElementById("ExcelStatus").innerHTML = "در حال پردازش اکسل لطفا منتظر بمانید";
    document.getElementById("submitExcel").style.display = "block";

    var btn = document.getElementById("submitExcel-btn");
    document.getElementById("submitExcel-btn").disabled = true;
    document.getElementById("submitExcel-btn").style.pointerEvents = "none";

    CheckExcelStatus();
}
function CheckExcelStatus() {
    excelInterval = setInterval(ExcelStatus, 10000);
}
function ExcelStatus() {
    debugger;
    const status = document.getElementById("ExcelStatus");
    status.classList.add("text-primary");
    fetch('/Home/ExcelStatus', { method: "GET" }).then(response => response.json())
        .then(data => {
            if (data.message === '') {
                status.innerHTML = "تعداد ردیف خوانده شده : " + data.row;
            }
            else {
                status.innerHTML = data.message;
                if (data.success) {
                    status.classList.add("text-success");
                }
                else {
                    status.classList.add("text-danger");
                }
                clearInterval(excelInterval);
                document.getElementById("submitExcel").style.display = "none";

                var btn = document.getElementById("submitExcel-btn");
                document.getElementById("submitExcel-btn").disabled = false;
                document.getElementById("submitExcel-btn").style.pointerEvents = "auto";

                document.getElementById("Excel").value = '';
            }

        });

}