import { PublicClientApplication, Configuration, DeviceCodeRequest } from '@azure/msal-node';
import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import * as dotenv from 'dotenv';
import * as fs from 'fs';

dotenv.config();

const CLIENT_ID = process.env.CLIENT_ID;
const TENANT_ID = process.env.TENANT_ID || "common";

if (!CLIENT_ID) {
    console.error("Missing CLIENT_ID in .env file.");
    process.exit(1);
}

const msalConfig: Configuration = {
    auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    }
};

const pca = new PublicClientApplication(msalConfig);

async function authenticate(): Promise<string> {
    const deviceCodeRequest: DeviceCodeRequest = {
        scopes: ["User.Read.All"],
        deviceCodeCallback: (response) => {
            console.log("\n========================================================");
            console.log(response.message);
            console.log("========================================================\n");
        },
    };

    try {
        console.log("Requesting device code...");
        const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
        if (response && response.accessToken) {
            console.log("Authentication successful!");
            return response.accessToken;
        } else {
            throw new Error("No access token returned.");
        }
    } catch (error) {
        console.error("Error during authentication:", error);
        process.exit(1);
    }
}

async function fetchAllUsers(client: Client) {
    let users: any[] = [];
    let nextLink = '/users?$select=id,displayName,jobTitle,mail,userPrincipalName,department,officeLocation&$expand=manager($select=id)';

    console.log("Fetching users...");
    while (nextLink) {
        try {
            const response = await client.api(nextLink).get();
            users = users.concat(response.value);
            nextLink = response['@odata.nextLink'];
        } catch (error) {
            console.error("Error fetching users:", error);
            break;
        }
    }
    return users;
}

async function fetchProfilePictureBase64(client: Client, userId: string): Promise<string | null> {
    try {
        const response = await client.api(`/users/${userId}/photo/$value`)
            .responseType(ResponseType.ARRAYBUFFER)
            .get();
        
        const buffer = Buffer.from(response);
        return `data:image/jpeg;base64,${buffer.toString('base64')}`;
    } catch (error: any) {
        // Many users might not have a profile picture, which returns 404
        return null;
    }
}

interface OrgNode {
    id: string;
    parentId: string | null;
    name: string;
    position: string;
    email: string;
    role: string;
    imageUrl: string | null;
    children: OrgNode[];
}

function buildTree(users: any[], photos: Record<string, string | null>): OrgNode[] {
    const nodeMap = new Map<string, OrgNode>();

    for (const user of users) {
        nodeMap.set(user.id, {
            id: user.id,
            parentId: user.manager?.id || null,
            name: user.displayName || "Unknown",
            position: user.jobTitle || "",
            email: user.mail || user.userPrincipalName || "",
            role: user.department || user.officeLocation || "",
            imageUrl: photos[user.id] || null,
            children: []
        });
    }

    const roots: OrgNode[] = [];

    for (const user of users) {
        const node = nodeMap.get(user.id)!;
        if (node.parentId && nodeMap.has(node.parentId)) {
            const parent = nodeMap.get(node.parentId);
            if (parent) {
                parent.children.push(node);
            } else {
                roots.push(node);
            }
        } else {
            roots.push(node);
        }
    }

    return roots;
}

async function main() {
    const accessToken = await authenticate();

    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    const users = await fetchAllUsers(client);
    console.log(`Fetched ${users.length} users.`);

    const photos: Record<string, string | null> = {};
    console.log("Fetching profile pictures...");
    
    const batchSize = 10;
    for (let i = 0; i < users.length; i += batchSize) {
        const batch = users.slice(i, i + batchSize);
        await Promise.all(batch.map(async (user) => {
            photos[user.id] = await fetchProfilePictureBase64(client, user.id);
        }));
        process.stdout.write(`\rProcessed ${Math.min(i + batchSize, users.length)} / ${users.length} photos`);
    }
    console.log("\nFinished fetching photos.");

    const tree = buildTree(users, photos);
    
    fs.writeFileSync('org_tree.json', JSON.stringify(tree, null, 2));
    console.log("Organization tree saved to org_tree.json!");
}

main().catch(console.error);
