Office.onReady(() => {
  // Categories for SLA and Escalation
  const categories = ["Estimate", "Complaint", "Query", "Feedback", "Spam", "Other"];
  
  // Populate SLA category dropdown
  const slaCategorySelect = document.getElementById("sla-category");
  categories.forEach(cat => {
    const option = document.createElement("option");
    option.value = cat;
    option.text = cat;
    slaCategorySelect.appendChild(option);
  });

  // Populate Escalation category dropdown
  const escalationCategorySelect = document.getElementById("escalation-category");
  categories.forEach(cat => {
    const option = document.createElement("option");
    option.value = cat;
    option.text = cat;
    escalationCategorySelect.appendChild(option);
  });
});

// Save SLA time
function saveSlaTime() {
  const selectedCategory = document.getElementById("sla-category").value;
  const slaTime = document.getElementById("sla-time").value;

  fetch("https://your-deployment-url/save_sla_time", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ category: selectedCategory, sla_time: slaTime })
  })
  .then(res => res.json())
  .then(data => {
    if (data.status === "success") {
      document.getElementById("status").innerText = "SLA Time saved successfully!";
      document.getElementById("status").style.color = "green";
    } else {
      document.getElementById("status").innerText = "Failed to save SLA Time.";
      document.getElementById("status").style.color = "red";
    }
  })
  .catch(err => {
    document.getElementById("status").innerText = "Error: " + err;
    document.getElementById("status").style.color = "red";
  });
}

// Save Escalation Person
function saveEscalationPerson() {
  const selectedCategory = document.getElementById("escalation-category").value;
  const escalationPerson = document.getElementById("escalation-person").value;

  fetch("https://your-deployment-url/save_escalation_person", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ category: selectedCategory, person: escalationPerson })
  })
  .then(res => res.json())
  .then(data => {
    if (data.status === "success") {
      document.getElementById("status").innerText = "Escalation Person saved successfully!";
      document.getElementById("status").style.color = "green";
    } else {
      document.getElementById("status").innerText = "Failed to save Escalation Person.";
      document.getElementById("status").style.color = "red";
    }
  })
  .catch(err => {
    document.getElementById("status").innerText = "Error: " + err;
    document.getElementById("status").style.color = "red";
  });
}

// Approve Suggested Reply
function approveReply() {
  const reply = document.getElementById("suggested-reply").value;

  fetch("https://your-deployment-url/approve_reply", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ reply: reply })
  })
  .then(res => res.json())
  .then(data => {
    if (data.status === "success") {
      document.getElementById("status").innerText = "Reply approved!";
      document.getElementById("status").style.color = "green";
    } else {
      document.getElementById("status").innerText = "Failed to approve reply.";
      document.getElementById("status").style.color = "red";
    }
  })
  .catch(err => {
    document.getElementById("status").innerText = "Error: " + err;
    document.getElementById("status").style.color = "red";
  });
}

// Edit Suggested Reply
function editReply() {
  const reply = document.getElementById("suggested-reply").value;
  // Open an editable area or log it for manual adjustments
  console.log("Edit reply:", reply);
}

// Decline Suggested Reply
function declineReply() {
  document.getElementById("status").innerText = "Reply declined.";
  document.getElementById("status").style.color = "red";
}
