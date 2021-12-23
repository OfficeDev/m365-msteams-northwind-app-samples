import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE } from './constants.js';

export async function getOrdersForEmployee(employeeId) {

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Orders?$filter=EmployeeID eq ${employeeId}&$top=10`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const data = await response.json();
    return data;
}

export async function getOrder(orderId) {

    const result = {};

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Orders(${orderId})?$expand=Customer,Employee`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const order = await response.json();

    result.orderId = order.OrderID;
    result.orderDate = order.OrderDate;
    result.requiredDate = order.RequiredDate;
    result.customerName = order.Customer.CompanyName;
    result.contactName = order.Customer.ContactName;
    result.contactTitle = order.Customer.ContactTitle;
    result.customerAddress = order.Customer.Address;
    result.customerCity = order.Customer.City;
    result.customerRegion = order.Customer.Region || "";
    result.customerPostalCode = order.Customer.PostalCode;
    result.customerPhone = order.Customer.Phone;
    result.customerCountry = order.Customer.Country;
    result.employeeName = `${order.Employee.FirstName} ${order.Employee.LastName}`;
    result.employeeEmail = `${order.Employee.LastName.toLowerCase()}@northwindtraders.com`;

    const response2 = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Order_Details?$filter=OrderID eq ${orderId}&$top=10&$expand=Product,Product/Category,Product/Supplier`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const details = await response2.json();

    result.details = details.value.map(lineItem => ({
        productName: lineItem.Product.ProductName,
        categoryName: lineItem.Product.Category.CategoryName,
        categoryPicture: lineItem.Product.Category.Picture.substring(104), // Remove Northwind-specific junk
        quantity: lineItem.Quantity,
        unitPrice: lineItem.UnitPrice,
        discount: lineItem.Discount,
        supplierName: lineItem.Product.Supplier.CompanyName,
        supplierCountry: lineItem.Product.Supplier.Country
    }));

    return result;
}