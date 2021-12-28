import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE } from './constants.js';

export async function getAllEmployees() {

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Employees/?$select=EmployeeID,FirstName,LastName`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });

    const employees = await response.json();
    return employees.value.map(employee => ({
        employeeId: employee.EmployeeID,
        firstName: employee.FirstName,
        lastName: employee.LastName
    }));
}

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

export async function getCategories() {

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Categories?$select=CategoryID,CategoryName,Description,Picture`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });

    const categories = await response.json();
    return categories.value.map(category => ({
        categoryId: category.CategoryID,
        displayName: category.CategoryName,
        description: category.Description,
        picture: category.Picture.substring(104), // Remove Northwind-specific junk
    }));
}

export async function getCategory(categoryId) {

    const result = {};

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Categories(${categoryId})?$select=CategoryID,CategoryName,Description,Picture`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const category = await response.json();

    result.categoryId = category.CategoryID;
    result.displayName = category.CategoryName;
    result.description = category.Description;
    result.picture = category.Picture.substring(104); // Remove Northwind-specific junk

    const response2 = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Products?$filter=CategoryID eq ${categoryId}&$expand=Supplier`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const details = await response2.json();

    result.products = details.value.map(product => ({
        productId: product.ProductID,
        productName: product.ProductName,
        quantityPerUnit: product.QuantityPerUnit,
        unitPrice: product.UnitPrice,
        unitsInStock: product.UnitsInStock,
        unitsOnOrder: product.UnitsOnOrder,
        reorderLevel: product.ReorderLevel,
        supplierName: product.Supplier.CompanyName,
        supplierCountry: product.Supplier.Country,
        discontinued: product.Discontinued
    }));

    return result;
}

export async function getProduct(productId) {
    
    const result = {};

    const response = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Products(${productId})?$expand=Category,Supplier`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const product = await response.json();
    result.productId = product.ProductID;
    result.productName = product.ProductName;
    result.quantityPerUnit = product.QuantityPerUnit;
    result.unitPrice = product.UnitPrice;
    result.unitsInStock = product.UnitsInStock;
    result.unitsOnOrder = product.UnitsOnOrder;
    result.reorderLevel = product.ReorderLevel;
    result.supplierName = product.Supplier.CompanyName;
    result.supplierCountry = product.Supplier.Country;
    result.discontinued = product.Discontinued;

    const response2 = await fetch(
        `${NORTHWIND_ODATA_SERVICE}/Order_Details?$filter=ProductID eq ${productId}&$expand=Order,Order/Customer,Order/Employee`,
        {
            "method": "GET",
            "headers": {
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
        });
    const details = await response2.json();

    result.orders = details.value.map(orderDetail => ({
        orderId: orderDetail.OrderID,
        customerId: orderDetail.Order.Customer.CustomerID,
        customerName: orderDetail.Order.Customer.CustomerName,
        customerAddress: `${orderDetail.Order.Customer.Address}, ${orderDetail.Order.Customer.City} ${orderDetail.Order.Customer.Region || ""}, ${orderDetail.Order.Customer.Country}`,
        employeeId: orderDetail.Order.EmployeeID,
        employeeName: `${orderDetail.Order.Employee.FirstName} ${orderDetail.Order.Employee.LastName}`,
        quantity: orderDetail.Quantity,
        unitPrice: orderDetail.UnitPrice,
        discount: orderDetail.Discount,
    }));

    return result;
}
