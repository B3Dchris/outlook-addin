Office.onReady(() => {
  Office.context.mailbox.item.from.getAsync((res) => {
    const email = res.value.emailAddress;
    document.getElementById("email-meta").innerText = "From: " + email;

    fetch("https://your-api.com/categories")
      .then(res => res.json())
      .then(data => {
        const select = document.getElementById("category-select");
        data.forEach(cat => {
          const option = document.createElement("option");
          option.value = cat.name;
          option.text = cat.name;
          select.appendChild(option);
        });
      });
  });
});

function saveCategory() {
  const selected = document.getElementById("category-select").value;
  fetch("https://your-api.com/assign-category", {
    method: "POST",
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({ category: selected })
  });
}
