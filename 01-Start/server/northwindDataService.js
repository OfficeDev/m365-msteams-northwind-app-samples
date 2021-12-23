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

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Orders(${orderId})`,
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

export async function getOrderDetails(orderId) {

    const response = await fetch(

        // To get this with the product and category info:
        // https://services.odata.org/V3/Northwind/Northwind.svc/Order_Details?$filter=OrderID eq 10265&$expand=Product,Product/Category
        `${NORTHWIND_ODATA_SERVICE}/Order_Details?$filter=OrderID eq ${orderId}&$top=10&$expand=Product,Product/Category,Product/Supplier`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const data = await response.json();

    return {
        productName: data.value[0].Product.ProductName,
        categoryName: data.value[0].Product.Category.CategoryName,
        categoryPicture: data.value[0].Product.Category.Picture,
        quantity: data.value[0].Quantity,
        unitPrice: data.value[0].UnitPrice,
        discount: data.value[0].Discount,
        supplierName: data.value[0].Product.Supplier.CompanyName,
        supplierCountry: data.value[0].Product.Supplier.Country
    }
}