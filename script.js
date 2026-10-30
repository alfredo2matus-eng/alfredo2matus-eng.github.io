
// ===================================
// 1. GET ELEMENTS FROM HTML
// ===================================

const form = document.getElementById("transaction-form");
const descriptionInput = document.getElementById("description");
const amountInput = document.getElementById("amount");
const typeInput = document.getElementById("type");
const monthInput = document.getElementById("month");

const incomeDisplay = document.getElementById("income");
const expensesDisplay = document.getElementById("expenses");
const balanceDisplay = document.getElementById("balance");
const transactionList = document.getElementById("transaction-list");

const exportButton = document.getElementById("export-button");
const chartCanvas = document.getElementById("budget-chart");

let budgetChart;


// ===================================
// 2. LOAD SAVED TRANSACTIONS
// ===================================

let transactions = JSON.parse(
    localStorage.getItem("myTransactions")
) || [];


// ===================================
// 3. SET THE CURRENT MONTH
// ===================================

const today = new Date();

const currentMonth =
    today.getFullYear() + "-" +
    String(today.getMonth() + 1).padStart(2, "0");

monthInput.value = currentMonth;


// ===================================
// 4. FORMAT MONEY
// ===================================

function formatMoney(amount) {
    return new Intl.NumberFormat("en-US", {
        style: "currency",
        currency: "USD"
    }).format(amount);
}


// ===================================
// 5. SAVE TRANSACTIONS
// ===================================

function saveTransactions() {
    localStorage.setItem(
        "myTransactions",
        JSON.stringify(transactions)
    );
}


// ===================================
// 6. DISPLAY TRANSACTIONS
// ===================================

function displayTransactions() {

    const selectedMonth = monthInput.value;

    // Show only transactions from selected month
    const monthlyTransactions = transactions.filter(
        transaction => transaction.month === selectedMonth
    );

    // Clear the old transaction list
    transactionList.innerHTML = "";

    let totalIncome = 0;
    let totalExpenses = 0;

    // Calculate income and expenses
    monthlyTransactions.forEach(transaction => {

        if (transaction.type === "income") {
            totalIncome += transaction.amount;
        } else {
            totalExpenses += transaction.amount;
        }

        // Create transaction row
        const item = document.createElement("li");
        item.className = "transaction-item";

        // Description
        const transactionText = document.createElement("span");
        transactionText.className = "transaction-description";
        transactionText.textContent = transaction.description;

        // Transaction details
        const details = document.createElement("div");
        details.className = "transaction-details";

        // Income or expense label
        const typeLabel = document.createElement("span");
        typeLabel.className = "transaction-type " + transaction.type;
        typeLabel.textContent =
            transaction.type === "income" ? "Income" : "Expense";

        // Amount
        const amountText = document.createElement("span");
        amountText.className = "transaction-amount " + transaction.type;

        const sign = transaction.type === "income" ? "+" : "-";

        amountText.textContent =
            sign + formatMoney(transaction.amount);

        // Delete button
        const deleteButton = document.createElement("button");
        deleteButton.textContent = "Delete";
        deleteButton.className = "delete-button";

        deleteButton.addEventListener("click", function() {
            deleteTransaction(transaction.id);
        });

        // Organize the row
        details.appendChild(typeLabel);
        details.appendChild(amountText);

        item.appendChild(transactionText);
        item.appendChild(details);
        item.appendChild(deleteButton);

        transactionList.appendChild(item);

    });

    // Calculate remaining balance
    const balance = totalIncome - totalExpenses;

    // Update the summary cards
    incomeDisplay.textContent = formatMoney(totalIncome);
    expensesDisplay.textContent = formatMoney(totalExpenses);
    balanceDisplay.textContent = formatMoney(balance);

    // Update the graph
    updateChart(totalIncome, totalExpenses);
}


// ===================================
// 7. UPDATE THE GRAPH
// ===================================

function updateChart(income, expenses) {

    if (budgetChart) {
        budgetChart.destroy();
    }

    budgetChart = new Chart(chartCanvas, {
        type: "bar",

        data: {
            labels: ["Income", "Expenses"],

            datasets: [{
                label: "Amount ($)",
                data: [income, expenses],
                backgroundColor: ["#18794e", "#c0392b"],
                borderRadius: 8
            }]
        },

        options: {
            responsive: true,

            scales: {
                y: {
                    beginAtZero: true
                }
            },

            plugins: {
                legend: {
                    display: false
                }
            }
        }
    });
}


// ===================================
// 8. ADD A NEW TRANSACTION
// ===================================

form.addEventListener("submit", function(event) {

    event.preventDefault();

    const description = descriptionInput.value.trim();
    const amount = Number(amountInput.value);
    const type = typeInput.value;

    // Check the information
    if (
        description === "" ||
        !Number.isFinite(amount) ||
        amount <= 0
    ) {
        alert("Please enter a description and a valid amount.");
        return;
    }

    // Create a transaction
    const newTransaction = {
        id: Date.now(),
        description: description,
        amount: amount,
        type: type,
        month: monthInput.value
    };

    // Add transaction to the array
    transactions.push(newTransaction);

    // Save and update the page
    saveTransactions();
    displayTransactions();

    // Clear the form
    form.reset();

});


// ===================================
// 9. DELETE A TRANSACTION
// ===================================

function deleteTransaction(id) {

    transactions = transactions.filter(
        transaction => transaction.id !== id
    );

    saveTransactions();
    displayTransactions();

}


// ===================================
// 10. CHANGE MONTH
// ===================================

monthInput.addEventListener("change", function() {
    displayTransactions();
});


// ===================================
// 11. EXPORT TO EXCEL
// ===================================

exportButton.addEventListener("click", function() {

    const selectedMonth = monthInput.value;

    if (!selectedMonth) {
        alert("Please select a month before exporting.");
        return;
    }

    const monthlyTransactions = transactions.filter(
        transaction => transaction.month === selectedMonth
    );

    let totalIncome = 0;
    let totalExpenses = 0;

    monthlyTransactions.forEach(transaction => {

        if (transaction.type === "income") {
            totalIncome += transaction.amount;
        } else {
            totalExpenses += transaction.amount;
        }

    });

    const balance = totalIncome - totalExpenses;

    // Create transaction worksheet
    const transactionData = monthlyTransactions.map(transaction => ({
        Description: transaction.description,
        Type: transaction.type,
        Amount: transaction.amount,
        Month: transaction.month
    }));

    const transactionSheet = XLSX.utils.json_to_sheet(
        transactionData
    );

    // Create summary worksheet
    const summaryData = [
        { Description: "Total Income", Amount: totalIncome },
        { Description: "Total Expenses", Amount: totalExpenses },
        { Description: "Balance", Amount: balance }
    ];

    const summarySheet = XLSX.utils.json_to_sheet(summaryData);

    // Create Excel workbook
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        summarySheet,
        "Monthly Summary"
    );

    XLSX.utils.book_append_sheet(
        workbook,
        transactionSheet,
        "Transactions"
    );

    // Download Excel file
    XLSX.writeFile(
        workbook,
        "My-Budget-" + selectedMonth + ".xlsx"
    );

});


// ===================================
// 12. START THE APP
// ===================================

displayTransactions();


