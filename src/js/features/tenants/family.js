/**
 * Family Table Feature Module
 * 
 * Handles all logic related to the family members table:
 * - Creating and managing table rows
 * - Syncing tenant details to the first row
 * - Collecting family member data
 */

/**
 * Syncs tenant details from form inputs to the first row of the family table
 * Creates the first row if it doesn't exist
 */
export function syncTenantToFamilyTable() {
    const tbody = document.querySelector("#familyTable tbody");
    if (!tbody) return;

    if (tbody.querySelectorAll("tr").length === 0) {
        const row = createFamilyRow();
        tbody.appendChild(row);
    }
}

/**
 * Creates a new family member table row
 * @param {Object} data - Family member data
 * @param {string} data.name - Member name
 * @param {string} data.relationship - Relationship to tenant
 * @param {string} data.occupation - Member occupation
 * @param {string} data.aadhaar - Aadhaar number
 * @param {string} data.address - Permanent address
 * @param {boolean} data.lockSelf - If true, makes the row read-only (for tenant's own row)
 * @returns {HTMLTableRowElement} The created table row
 */
export function createFamilyRow(data = {}) {
    const tr = document.createElement("tr");

    // Name column
    const nameTd = document.createElement("td");
    nameTd.className = "border px-2 py-1";
    const nameInput = document.createElement("input");
    nameInput.type = "text";
    nameInput.className = "w-full text-xs border rounded px-1 py-0.5";
    nameInput.value = data.name || "";
    if (data.lockSelf) nameInput.readOnly = true;
    nameTd.appendChild(nameInput);

    // Relationship column
    const relTd = document.createElement("td");
    relTd.className = "border px-2 py-1";
    const relInput = document.createElement("input");
    relInput.type = "text";
    relInput.className = "w-full text-xs border rounded px-1 py-0.5";
    relInput.value = data.relationship || "";
    if (data.lockSelf) relInput.readOnly = true;
    relTd.appendChild(relInput);

    // Occupation column
    const occTd = document.createElement("td");
    occTd.className = "border px-2 py-1";
    const occInput = document.createElement("input");
    occInput.type = "text";
    occInput.className = "w-full text-xs border rounded px-1 py-0.5";
    occInput.value = data.occupation || "";
    if (data.lockSelf) occInput.readOnly = true;
    occTd.appendChild(occInput);

    // Aadhaar column
    const aadTd = document.createElement("td");
    aadTd.className = "border px-2 py-1";
    const aadInput = document.createElement("input");
    aadInput.type = "text";
    aadInput.className = "w-full text-xs border rounded px-1 py-0.5";
    aadInput.value = data.aadhaar || "";
    if (data.lockSelf) aadInput.readOnly = true;
    aadTd.appendChild(aadInput);

    // Address column
    const addrTd = document.createElement("td");
    addrTd.className = "border px-2 py-1";
    const addrInput = document.createElement("textarea");
    addrInput.rows = 1;
    addrInput.className = "w-full text-xs border rounded px-1 py-0.5 resize-y";
    addrInput.value = data.address || "";
    if (data.lockSelf) addrInput.readOnly = true;
    addrTd.appendChild(addrInput);

    // Delete button column
    const delTd = document.createElement("td");
    delTd.className = "border px-2 py-1 text-center";
    const delBtn = document.createElement("button");
    delBtn.type = "button";
    delBtn.textContent = "×";
    delBtn.className =
        "text-xs px-2 py-0.5 rounded border border-red-400 text-red-600 hover:bg-red-50";
    if (data.lockSelf) {
        delBtn.disabled = true;
        delBtn.classList.add("opacity-40", "cursor-not-allowed");
    } else {
        delBtn.addEventListener("click", () => {
            tr.remove();
        });
    }
    delTd.appendChild(delBtn);

    tr.appendChild(nameTd);
    tr.appendChild(relTd);
    tr.appendChild(occTd);
    tr.appendChild(aadTd);
    tr.appendChild(addrTd);
    tr.appendChild(delTd);

    return tr;
}

/**
 * Collects all family member data from the table
 * @returns {Array<Object>} Array of family member objects
 */
export function getFamilyMembersFromTable() {
    const rows = Array.from(document.querySelectorAll("#familyTable tbody tr"));
    return rows
        .map((tr) => {
            const inputs = tr.querySelectorAll("input, textarea");
            const member = {
                name: inputs[0].value.trim(),
                relationship: inputs[1].value.trim(),
                occupation: inputs[2].value.trim(),
                aadhaar: inputs[3].value.trim(),
                address: inputs[4].value.trim(),
            };

            if (
                !member.name &&
                !member.relationship &&
                !member.occupation &&
                !member.aadhaar &&
                !member.address
            ) {
                return null;
            }

            return member;
        })
        .filter(Boolean);
}

/**
 * Initializes the family table with the tenant's row
 */
export function initFamilyTable() {
    syncTenantToFamilyTable();
}

/**
 * Rebuilds the family table based on provided member data
 * @param {Array<Object>} members - Family members to populate
 */
export function setFamilyMembersInTable(members = []) {
    const tbody = document.querySelector("#familyTable tbody");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (!members.length) {
        syncTenantToFamilyTable();
        return;
    }

    members.forEach((member) => {
        const row = createFamilyRow({
            name: member.name,
            relationship: member.relationship,
            occupation: member.occupation,
            aadhaar: member.aadhaar,
            address: member.address,
        });
        tbody.appendChild(row);
    });
}
