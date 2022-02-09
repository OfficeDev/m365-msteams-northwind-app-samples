## Lab A06: Extend teams application with Messaging Extension

In this lab you will begin with the source code folder `A05-ConfigurableTab`, make changes as per the steps below to achieve what is in the source code folder `A06-MessagingExtension`.
See project structures before and after.

<table>
<tr>
<th>A05-ConfigurableTab - Before </th>
<th>A06-MessagingExtension - After</th>
</tr>
<tr>
<td style=" vertical-align: top;">
<pre>
A05-ConfigurableTab
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
    │       └── userPanel.js
    ├── modules
    │   └── env.js
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── orderDetail.html
    │   └── orderDetail.js
    │   └── privacy.html
    │   └── productDetail.html
    │   └── productDetail.js
    │   └── tabConfig.html
    │   └── tabConfig.js
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── teamstyle.css
    ├── manifest
    │   └── <span style="color:red">makePackage.js</span>
    │   └── <span style="color:red">manifest.template.json</span>
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── constants.js
    │   └── <span style="color:red">identityService.js</span>
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── <span style="color:red">server.js</span>
    ├── <span style="color:red">.env_Sample</span>
    ├── .gitignore
    ├── <span style="color:red">package.json</span>
    ├── README.md
</pre>
</td>
<td>
<pre>
A06-MessagingExtension
    ├── client
    │   ├── components
    │       ├── navigation.js
    │   └── identity
    │       ├── identityClient.js
    │       └── userPanel.js
    ├── <span style="color:purple">images</span>
    │   └── <span style="color:purple">1.PNG</span>
    │   └── <span style="color:purple">2.PNG</span>
    │   └── <span style="color:purple">3.PNG</span>
    │   └── <span style="color:purple">4.PNG</span>
    │   └── <span style="color:purple">5.PNG</span>
    │   └── <span style="color:purple">6.PNG</span>
    │   └── <span style="color:purple">7.PNG</span>
    │   └── <span style="color:purple">8.PNG</span>
    │   └── <span style="color:purple">9.PNG</span>
    ├── modules
    │   └── env.js
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── teamsHelpers.js
    ├── pages
    │   └── categories.html
    │   └── categories.js
    │   └── categoryDetails.html
    │   └── categoryDetails.js
    │   └── myOrders.html
    │   └── orderDetail.html
    │   └── orderDetail.js
    │   └── privacy.html
    │   └── productDetail.html
    │   └── productDetail.js
    │   └── tabConfig.html
    │   └── tabConfig.js
    │   └── termsofuse.html
    ├── index.html
    ├── index.js
    ├── northwind.css
    ├── teamstyle.css
    ├── manifest
    │   └── <span style="color:red">makePackage.js</span>
    │   └── <span style="color:red">manifest.template.json</span>
    │   └── northwind32.png
    │   └── northwind192.png
    ├── server
    │   └── <span style="color:purple">cards</span>
    │       └── <span style="color:purple">errorCard.js</span>
    │       └── <span style="color:purple">productCard.js</span>
    │       └── <span style="color:purple">stockUpdateSuccess.js</span>
    │   └── <span style="color:purple">bot.js</span>
    │   └── constants.js
    │   └── <span style="color:red">identityService.js</span>
    │   └── <span style="color:red">northwindDataService.js</span>
    │   └── <span style="color:red">server.js</span>
    ├── <span style="color:red">.env_Sample</span>
    ├── .gitignore
    ├── <span style="color:red">package.json</span>
    ├── README.md
</pre>
</td>
</tr>
</table>


### Exercise 1: Bot registration

#### Step 1: Create Bot framework registration resource

#### Step 2: Enable the Teams Channel

#### Step 3: Run ngrok 

#### Step 4: Update the bot registration configuration

#### Step 5: Exercise 2: Source code changes


### Exercise 2: Code changes

#### Step 1: Add new files

#### Step 2: Update existing files


### Exercise 3: Test the changes


### Known issues

### References



