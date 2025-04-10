Office.onReady(() => {
  const categories = ["Billing", "Support", "Urgent", "Low Priority"];
  const select = document.getElementById("category-select");
  categories.forEach(cat => {
    const option = document.createElement("option");
    option.value = cat;
    option.text = cat;
    select.appendChild(option);
  });
});

function saveCategory() {
  const selected = document.getElementById("category-select").value;
  document.getElementById("status").innerText = "Saved category: " + selected;
}
