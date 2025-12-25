import assert from "node:assert/strict";
import { generateNoGrnValue } from "../src/js/utils/grn.js";
import { resolveDraftHydration } from "../src/js/features/tenants/draftHydration.js";

function testGenerateNoGrnValue() {
    const fixedZero = generateNoGrnValue(() => 0);
    assert.equal(fixedZero, "NoGRN00000");

    const fixedMax = generateNoGrnValue(() => 0.99999);
    assert.equal(fixedMax, "NoGRN99999");

    const sample = generateNoGrnValue(() => 0.12345);
    assert.match(sample, /^NoGRN\d{5}$/);
}

function testResolveDraftHydrationUsesCache() {
    const values = {
        landlord_selector: "l1",
        unit_selector: "u1",
        grnNoCheckbox: true,
        grn_number: "",
    };
    const landlords = [{ landlord_id: "l1", name: "Landlord A", aadhaar: "123", address: "Addr" }];
    const units = [{ unit_id: "u1", wing: "A", unit_number: "101", floor: "1", direction: "N", meter_number: "M1" }];

    const result = resolveDraftHydration(values, {
        landlords,
        units,
        generateNoGrnValue: () => "NoGRN99999",
    });

    assert.equal(result.landlordFields.name, "Landlord A");
    assert.equal(result.unitFields.wing, "A");
    assert.equal(result.noGrnChecked, true);
    assert.equal(result.grnValue, "NoGRN99999");
}

function testResolveDraftHydrationFallbacks() {
    const values = {
        Landlord_name: "Fallback Name",
        landlord_aadhar: "999",
        landlord_address: "Fallback Addr",
        wing: "B",
        unit_number_display: "202",
        floor_of_building: "2",
        direction_build: "E",
        meter_number: "M2",
        grnNoCheckbox: false,
        grn_number: "GRN-1",
    };

    const result = resolveDraftHydration(values, { landlords: [], units: [] });

    assert.equal(result.landlordFields.name, "Fallback Name");
    assert.equal(result.unitFields.wing, "B");
    assert.equal(result.noGrnChecked, false);
    assert.equal(result.grnValue, "GRN-1");
}

function run() {
    testGenerateNoGrnValue();
    testResolveDraftHydrationUsesCache();
    testResolveDraftHydrationFallbacks();
    console.log("All tests passed.");
}

run();
