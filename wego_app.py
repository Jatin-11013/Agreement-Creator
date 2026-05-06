import streamlit as st

st.set_page_config(page_title="Meta Fees Calculator", layout="centered")
st.title("🧮 Meta Fees Calculator")
st.write("Wego & TripSaverz – Operation Team Tool")

# -------- INPUTS --------
meta_partner = st.selectbox("Meta Partner", ["Wego", "TripSaverz"])
flight_type = st.selectbox("Flight Type", ["Domestic", "International"])
booking_amount = st.number_input(
    "Booking Amount (₹)",
    min_value=0.0,
    step=100.0
)
pax_count = st.number_input(
    "Passenger Count",
    min_value=1,
    step=1
)

# -------- FEES LOGIC --------
def calculate_meta_fee(meta, flight, amount, pax):
    if meta == "Wego":
        if flight == "Domestic":
            return 200 if pax <= 2 else 300
        elif flight == "International":
            return 400 if amount <= 30000 else 600

    elif meta == "TripSaverz":
        if flight == "Domestic":
            return amount * 0.01
        elif flight == "International":
            return amount * 0.015

    return 0

# -------- CALCULATE BUTTON --------
st.divider()

if st.button("Calculate Meta Fees"):
    meta_fee = calculate_meta_fee(
        meta_partner,
        flight_type,
        booking_amount,
        pax_count
    )

    st.subheader("💰 Meta Fees")
    st.write(f"₹ {meta_fee:.2f}")
