"""
FedEx Shipping API Integration Module

Handles OAuth2 authentication, shipment creation, and label generation
for FedEx Ground shipping.
"""

import os
import requests
import base64
from datetime import datetime, timedelta
from pathlib import Path


# FedEx API endpoints
FEDEX_AUTH_URL = "https://apis.fedex.com/oauth/token"
FEDEX_SHIP_URL = "https://apis.fedex.com/ship/v1/shipments"

# Sandbox endpoints for testing
FEDEX_AUTH_URL_SANDBOX = "https://apis-sandbox.fedex.com/oauth/token"
FEDEX_SHIP_URL_SANDBOX = "https://apis-sandbox.fedex.com/ship/v1/shipments"

# Token cache
_token_cache = {
    "token": None,
    "expires_at": None
}


def get_fedex_token(api_key, secret_key, use_sandbox=False):
    """
    Authenticate with FedEx OAuth2 and get access token.

    Args:
        api_key: FedEx API Key (Client ID)
        secret_key: FedEx Secret Key (Client Secret)
        use_sandbox: Use sandbox environment for testing

    Returns:
        Access token string, or None if authentication fails
    """
    global _token_cache

    # Check if we have a valid cached token
    if _token_cache["token"] and _token_cache["expires_at"]:
        if datetime.now() < _token_cache["expires_at"]:
            return _token_cache["token"]

    auth_url = FEDEX_AUTH_URL_SANDBOX if use_sandbox else FEDEX_AUTH_URL

    try:
        response = requests.post(
            auth_url,
            data={
                "grant_type": "client_credentials",
                "client_id": api_key,
                "client_secret": secret_key
            },
            headers={
                "Content-Type": "application/x-www-form-urlencoded"
            },
            timeout=30
        )

        if response.status_code == 200:
            data = response.json()
            token = data.get("access_token")
            expires_in = data.get("expires_in", 3600)  # Default 1 hour

            # Cache the token with expiration (subtract 5 min for safety margin)
            _token_cache["token"] = token
            _token_cache["expires_at"] = datetime.now() + timedelta(seconds=expires_in - 300)

            print(f"[FedEx] Successfully authenticated")
            return token
        else:
            print(f"[FedEx] Authentication failed: {response.status_code}")
            print(f"[FedEx] Response: {response.text}")
            return None

    except requests.exceptions.RequestException as e:
        print(f"[FedEx] Authentication error: {e}")
        return None


def create_shipment(token, account_number, shipper, recipient, package_details, use_sandbox=False):
    """
    Create a FedEx shipment and get shipping label.

    Args:
        token: FedEx OAuth access token
        account_number: FedEx account number for billing
        shipper: Dict with shipper address info
        recipient: Dict with recipient address info
        package_details: Dict with package weight, dimensions
        use_sandbox: Use sandbox environment for testing

    Returns:
        Dict with tracking_number and label_data (base64 PDF), or None if failed
    """
    ship_url = FEDEX_SHIP_URL_SANDBOX if use_sandbox else FEDEX_SHIP_URL

    # Build the shipment request
    shipment_request = {
        "labelResponseOptions": "LABEL",
        "requestedShipment": {
            "shipper": {
                "contact": {
                    "personName": shipper.get("contact_name", shipper.get("company", "")),
                    "phoneNumber": shipper.get("phone", ""),
                    "companyName": shipper.get("company", "")
                },
                "address": {
                    "streetLines": [shipper.get("street", "")],
                    "city": shipper.get("city", ""),
                    "stateOrProvinceCode": shipper.get("state", ""),
                    "postalCode": shipper.get("zip", ""),
                    "countryCode": "US"
                }
            },
            "recipients": [{
                "contact": {
                    "personName": recipient.get("name", ""),
                    "phoneNumber": recipient.get("phone", "")
                },
                "address": {
                    "streetLines": [
                        recipient.get("address1", ""),
                        recipient.get("address2", "")
                    ] if recipient.get("address2") else [recipient.get("address1", "")],
                    "city": recipient.get("city", ""),
                    "stateOrProvinceCode": recipient.get("state", ""),
                    "postalCode": recipient.get("zip", ""),
                    "countryCode": "US",
                    "residential": True
                }
            }],
            "shipDatestamp": datetime.now().strftime("%Y-%m-%d"),
            "serviceType": "FEDEX_GROUND",
            "packagingType": "YOUR_PACKAGING",
            "pickupType": "USE_SCHEDULED_PICKUP",
            "blockInsightVisibility": False,
            "shippingChargesPayment": {
                "paymentType": "SENDER",
                "payor": {
                    "responsibleParty": {
                        "accountNumber": {
                            "value": account_number
                        }
                    }
                }
            },
            "labelSpecification": {
                "imageType": "PDF",
                "labelStockType": "PAPER_4X6"
            },
            "requestedPackageLineItems": [{
                "weight": {
                    "units": "LB",
                    "value": package_details.get("weight", 1.0)
                }
            }]
        },
        "accountNumber": {
            "value": account_number
        }
    }

    # Add dimensions if provided
    if package_details.get("length") and package_details.get("width") and package_details.get("height"):
        shipment_request["requestedShipment"]["requestedPackageLineItems"][0]["dimensions"] = {
            "length": package_details.get("length"),
            "width": package_details.get("width"),
            "height": package_details.get("height"),
            "units": "IN"
        }

    try:
        response = requests.post(
            ship_url,
            json=shipment_request,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}",
                "X-locale": "en_US"
            },
            timeout=60
        )

        if response.status_code == 200:
            data = response.json()

            # Extract tracking number and label
            output = data.get("output", {})
            transaction_shipments = output.get("transactionShipments", [])

            if transaction_shipments:
                shipment = transaction_shipments[0]
                tracking_number = shipment.get("masterTrackingNumber", "")

                # Get label data
                piece_responses = shipment.get("pieceResponses", [])
                label_data = None

                if piece_responses:
                    package_documents = piece_responses[0].get("packageDocuments", [])
                    if package_documents:
                        label_data = package_documents[0].get("encodedLabel", "")

                print(f"[FedEx] Shipment created - Tracking: {tracking_number}")

                return {
                    "tracking_number": tracking_number,
                    "label_data": label_data
                }

        print(f"[FedEx] Shipment creation failed: {response.status_code}")
        print(f"[FedEx] Response: {response.text}")
        return None

    except requests.exceptions.RequestException as e:
        print(f"[FedEx] Shipment creation error: {e}")
        return None


def save_label_pdf(label_data, output_path):
    """
    Save base64-encoded label data as PDF file.

    Args:
        label_data: Base64-encoded PDF string
        output_path: Path to save the PDF file

    Returns:
        True if saved successfully, False otherwise
    """
    try:
        pdf_bytes = base64.b64decode(label_data)

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with open(output_path, 'wb') as f:
            f.write(pdf_bytes)

        print(f"[FedEx] Label saved: {output_path}")
        return True

    except Exception as e:
        print(f"[FedEx] Error saving label: {e}")
        return False


def get_shipping_label(order, ship_from_location, config):
    """
    Main function to generate FedEx shipping label for an order.

    Args:
        order: Order dict from Supabase
        ship_from_location: "Yakima" or "Toppenish"
        config: Application config dict with FedEx credentials

    Returns:
        Dict with tracking_number and label_path, or None if failed
    """
    # Get FedEx credentials from config
    api_key = config.get("fedex_api_key")
    secret_key = config.get("fedex_secret_key")
    account_number = config.get("fedex_account_number")
    shipper_addresses = config.get("shipper_addresses", {})
    use_sandbox = config.get("fedex_use_sandbox", False)

    if not all([api_key, secret_key, account_number]):
        print("[FedEx] Missing FedEx credentials in config")
        return None

    # Get shipper address for the location
    shipper = shipper_addresses.get(ship_from_location)
    if not shipper:
        print(f"[FedEx] No shipper address configured for {ship_from_location}")
        return None

    # Build recipient from order shipping address
    shipping_address = order.get("customer_shipping_address", {})
    if not shipping_address:
        print("[FedEx] No shipping address in order")
        return None

    recipient = {
        "name": f"{order.get('customer_first_name', '')} {order.get('customer_last_name', '')}".strip(),
        "phone": order.get("customer_phone", ""),
        "address1": shipping_address.get("address1", ""),
        "address2": shipping_address.get("address2", ""),
        "city": shipping_address.get("city", ""),
        "state": shipping_address.get("state", ""),
        "zip": shipping_address.get("zipCode", "")
    }

    # Calculate package weight from order items
    # Default to 1 lb if not specified
    total_weight = 0
    for item in order.get("items", []):
        item_weight = item.get("weight", 0.5)  # Default 0.5 lb per item
        quantity = item.get("quantity", 1)
        total_weight += item_weight * quantity

    if total_weight < 1:
        total_weight = 1.0  # Minimum 1 lb

    package_details = {
        "weight": round(total_weight, 1)
    }

    # Authenticate
    token = get_fedex_token(api_key, secret_key, use_sandbox)
    if not token:
        return None

    # Create shipment
    result = create_shipment(
        token=token,
        account_number=account_number,
        shipper=shipper,
        recipient=recipient,
        package_details=package_details,
        use_sandbox=use_sandbox
    )

    if not result:
        return None

    # Save label PDF
    from inventory_sync import PDF_OUTPUT_DIR
    label_filename = f"shipping_label_{order['order_number']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    label_path = PDF_OUTPUT_DIR / label_filename

    if result.get("label_data"):
        if save_label_pdf(result["label_data"], label_path):
            return {
                "tracking_number": result["tracking_number"],
                "label_path": str(label_path)
            }

    return {
        "tracking_number": result["tracking_number"],
        "label_path": None
    }


def has_shipping_items(order):
    """
    Check if an order contains any items with shipping fulfillment.

    Args:
        order: Order dict from Supabase

    Returns:
        True if order has shipping items, False otherwise
    """
    for item in order.get("items", []):
        fulfillment = item.get("fulfillment", {})
        if fulfillment.get("method") == "shipping":
            return True
    return False


def get_ship_from_location(order, default_location="Toppenish"):
    """
    Determine which location to ship from based on order data.

    Args:
        order: Order dict from Supabase
        default_location: Default location if not specified

    Returns:
        "Yakima" or "Toppenish"
    """
    # Check order_location field
    order_location = order.get("order_location", "").lower()

    if order_location == "yakima":
        return "Yakima"
    elif order_location == "toppenish":
        return "Toppenish"

    # Check if any shipping items specify a location
    for item in order.get("items", []):
        fulfillment = item.get("fulfillment", {})
        if fulfillment.get("method") == "shipping":
            location = fulfillment.get("location") or fulfillment.get("shipFrom")
            if location:
                if str(location).lower() == "yakima" or location == 1:
                    return "Yakima"
                elif str(location).lower() == "toppenish" or location == 2:
                    return "Toppenish"

    return default_location
