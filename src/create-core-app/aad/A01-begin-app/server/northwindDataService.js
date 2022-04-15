import fetch from 'node-fetch';
import { NORTHWIND_ODATA_SERVICE, EMAIL_DOMAIN } from './constants.js';
import { dbService } from '../northwindDB/dbService.js';

const db = new dbService();     // Singleton service for Northwind DB

const employeeCache = {};
export async function getEmployee(employeeId) {

    if (employeeCache[employeeId]) return employeeCache[employeeId];

    const employees = await db.getTable("Employees", "EmployeeID");
    const orders = await db.getTable("Orders", "OrderID");
    const customers = await db.getTable("Customers", "CustomerID");

    const result = {};

    const employee = employees.item(employeeId);

    result.id = employee.EmployeeID;
    result.displayName = `${employee.FirstName} ${employee.LastName}`;
    result.mail = `${employee.FirstName}@${EMAIL_DOMAIN}`;
    result.photo = employee.Photo.substring(104); // Trim Northwind-specific junk
    result.jobTitle = employee.Title;
    result.city = `${employee.City}, ${employee.Region || ''} ${employee.Country}`;

    const employeeOrders = orders.data.filter((order) => order.EmployeeID === result.id);

    result.orders = employeeOrders.map((order) => {
        const customer = customers.item(order.CustomerID);
        return ({
            orderId: order.OrderID,
            orderDate: order.OrderDate,
            customerId: order.CustomerID,
            customerName: customer.CompanyName,
            customerContact: customer.ContactName,
            customerPhone: customer.Phone,
            shipName: order.ShipName,
            shipAddress: order.ShipAddress,
            shipCity: order.shipCity,
            shipRegion: order.ShipRegion,
            shipPostalCode: order.shipPostalCode,
            shipCountry: order.shipCountry
        });
    });

    employeeCache[employeeId] = result;
    return result;
}

const orderCache = {}
export async function getOrder(orderId) {

    if (orderCache[orderId]) return orderCache[orderId];

    const result = {};

    const orders = await db.getTable("Orders", "OrderID");
    const customers = await db.getTable("Customers", "CustomerID");
    const employees = await db.getTable("Employees", "EmployeeID");

    const order = orders.item(orderId);
    const customer = customers.item(order.CustomerID);
    const employee = employees.item(order.EmployeeID);

    result.orderId = orderId;
    result.orderDate = order.OrderDate;
    result.requiredDate = order.RequiredDate;
    result.customerName = customer.CompanyName;
    result.contactName = customer.ContactName;
    result.contactTitle = customer.ContactTitle;
    result.customerAddress = customer.Address;
    result.customerCity = customer.City;
    result.customerRegion = customer.Region || "";
    result.customerPostalCode = customer.PostalCode;
    result.customerPhone = customer.Phone;
    result.customerCountry = customer.Country;
    result.employeeId = employee.EmployeeID;
    result.employeeName = `${employee.FirstName} ${employee.LastName}`;
    result.employeeEmail = `${employee.LastName.toLowerCase()}@northwindtraders.com`;
    result.employeeTitle = `${employee.Title}`;

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
        productId: lineItem.ProductID,
        productName: lineItem.Product.ProductName,
        categoryName: lineItem.Product.Category.CategoryName,
        categoryPicture: lineItem.Product.Category.Picture.substring(104), // Remove Northwind-specific junk
        quantity: lineItem.Quantity,
        unitPrice: lineItem.UnitPrice,
        discount: lineItem.Discount,
        supplierName: lineItem.Product.Supplier.CompanyName,
        supplierCountry: lineItem.Product.Supplier.Country
    }));

    orderCache[orderId] = result;
    return result;
}

const categoriesCache = {};
export async function getCategories() {

    if (categoriesCache.data) return categoriesCache.data;

    const categories = await db.getTable("Categories", "CategoryID");

    const result = categories.data.map(category => ({
        categoryId: category.CategoryID,
        displayName: category.CategoryName,
        description: category.Description,
        picture: category.Picture.substring(104), // Remove Northwind-specific junk
    }));
    categoriesCache.data = result;
    return result;
}

const categoryCache = {};
export async function getCategory(categoryId) {

    if (categoryCache[categoryId]) return categoryCache[categoryId];

    const result = {};

    const categories = await db.getTable("Categories", "CategoryID");
    const category = categories.item(categoryId);

    result.categoryId = category.CategoryID;
    result.displayName = category.CategoryName;
    result.description = category.Description;
    result.picture = category.Picture.substring(104); // Remove Northwind-specific junk

    const products = await db.getTable("Products", "ProductID");
    const suppliers = await db.getTable("Suppliers", "SupplierID");
    const productsInCategory = products.data.filter(product => product.CategoryID == categoryId);

    result.products = productsInCategory.map((product) => {
        const supplier = suppliers.item([product.SupplierID]);
        return ({
            productId: product.ProductID,
            productName: product.ProductName,
            quantityPerUnit: product.QuantityPerUnit,
            unitPrice: product.UnitPrice,
            unitsInStock: product.UnitsInStock,
            unitsOnOrder: product.UnitsOnOrder,
            reorderLevel: product.ReorderLevel,
            supplierName: supplier.CompanyName,
            supplierCountry: supplier.Country,
            discontinued: product.Discontinued
        })
    }).sort((a, b) => a.productName.localeCompare(b.productName));

    categoryCache[categoryId] = result;
    return result;
}

const productCache = {};
export async function getProduct(productId) {

    if (productCache[productId]) return productCache[productId];

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
    result.categoryId = product.CategoryID;
    result.categoryName = product.Category.CategoryName;
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
        orderDate: orderDetail.Order.OrderDate,
        customerId: orderDetail.Order.Customer.CustomerID,
        customerName: orderDetail.Order.Customer.CompanyName,
        customerAddress: `${orderDetail.Order.Customer.Address}, ${orderDetail.Order.Customer.City} ${orderDetail.Order.Customer.Region || ""}, ${orderDetail.Order.Customer.Country}`,
        employeeId: orderDetail.Order.EmployeeID,
        employeeName: `${orderDetail.Order.Employee.FirstName} ${orderDetail.Order.Employee.LastName}`,
        quantity: orderDetail.Quantity,
        unitPrice: orderDetail.UnitPrice,
        discount: orderDetail.Discount,
    }));

    productCache[productId] = result;
    return result;
}
