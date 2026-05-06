def format_naira(amount):
    if amount is None:
        return "₦0.00"
    return f"₦{float(amount):,.2f}"
