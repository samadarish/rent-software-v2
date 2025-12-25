const DRAFT_EXCLUDED_FIELDS = new Set(["rent_amount_words", "secu_amount_words"]);

export function pruneDraftValues(values = {}) {
    if (!values || typeof values !== "object") return {};
    const pruned = {};
    Object.entries(values).forEach(([key, value]) => {
        if (DRAFT_EXCLUDED_FIELDS.has(key)) return;
        pruned[key] = value;
    });
    return pruned;
}

export function resolveDraftHydration(values = {}, options = {}) {
    const landlords = Array.isArray(options.landlords) ? options.landlords : [];
    const units = Array.isArray(options.units) ? options.units : [];

    const landlordId = values.landlord_selector || values.landlord_id || "";
    const matchedLandlord = landlords.find((l) => l.landlord_id === landlordId) || null;
    const landlordFields = matchedLandlord
        ? {
              name: matchedLandlord.name || "",
              aadhaar: matchedLandlord.aadhaar || "",
              address: matchedLandlord.address || "",
          }
        : {
              name: values.Landlord_name || "",
              aadhaar: values.landlord_aadhar || "",
              address: values.landlord_address || "",
          };

    const unitId = values.unit_selector || values.unit_id || "";
    const matchedUnit = units.find((u) => u.unit_id === unitId) || null;
    const unitFields = matchedUnit
        ? {
              wing: matchedUnit.wing || "",
              unitNumber: matchedUnit.unit_number || "",
              floor: matchedUnit.floor || "",
              direction: matchedUnit.direction || "",
              meter: matchedUnit.meter_number || "",
          }
        : {
              wing: values.wing || "",
              unitNumber: values.unit_number_display || values.unit_number || "",
              floor: values.floor_of_building || "",
              direction: values.direction_build || "",
              meter: values.meter_number || "",
          };

    const noGrnChecked = Boolean(values.grnNoCheckbox);
    let grnValue = values.grn_number || "";
    if (noGrnChecked && !grnValue && typeof options.generateNoGrnValue === "function") {
        grnValue = options.generateNoGrnValue();
    }

    return {
        landlordId,
        landlordFields,
        unitId,
        unitFields,
        noGrnChecked,
        grnValue,
    };
}
