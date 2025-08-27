import { useState } from "react";
import { Search, Filter, Copy, ExternalLink, Check } from "lucide-react";
import CliveDownloader from "@/components/CliveDownloader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Checkbox } from "@/components/ui/checkbox";
import { cn } from "@/lib/utils";

interface GraphEndpoint {
  id: string;
  name: string;
  url: string;
  method: string;
  category: string;
  scopes: string[];
  description: string;
  version: "v1.0" | "beta";
}

const endpoints: GraphEndpoint[] = [
  {
    id: "users-list",
    name: "List Users",
    url: "https://graph.microsoft.com/v1.0/users",
    method: "GET",
    category: "Users",
    scopes: ["User.Read.All", "User.ReadWrite.All", "Directory.Read.All"],
    description: "Retrieve a list of user objects",
    version: "v1.0",
  },
  {
    id: "user-get",
    name: "Get User",
    url: "https://graph.microsoft.com/v1.0/users/{id}",
    method: "GET",
    category: "Users",
    scopes: ["User.Read", "User.Read.All"],
    description: "Retrieve properties and relationships of user object",
    version: "v1.0",
  },
  {
    id: "groups-list",
    name: "List Groups",
    url: "https://graph.microsoft.com/v1.0/groups",
    method: "GET",
    category: "Groups",
    scopes: ["Group.Read.All", "Group.ReadWrite.All", "Directory.Read.All"],
    description: "List all groups in an organization",
    version: "v1.0",
  },
  {
    id: "mail-list",
    name: "List Messages",
    url: "https://graph.microsoft.com/v1.0/me/messages",
    method: "GET",
    category: "Mail",
    scopes: ["Mail.Read", "Mail.ReadWrite"],
    description: "Get the messages in the signed-in user's mailbox",
    version: "v1.0",
  },
  {
    id: "calendar-events",
    name: "List Events",
    url: "https://graph.microsoft.com/v1.0/me/events",
    method: "GET",
    category: "Calendar",
    scopes: ["Calendars.Read", "Calendars.ReadWrite"],
    description: "Get events from the user's primary calendar",
    version: "v1.0",
  },
  {
    id: "files-list",
    name: "List Drive Items",
    url: "https://graph.microsoft.com/v1.0/me/drive/root/children",
    method: "GET",
    category: "Files",
    scopes: ["Files.Read", "Files.ReadWrite", "Files.Read.All"],
    description: "List the children of a driveItem",
    version: "v1.0",
  },
  {
    id: "applications-list",
    name: "List Applications",
    url: "https://graph.microsoft.com/v1.0/applications",
    method: "GET",
    category: "Applications",
    scopes: ["Application.Read.All", "Application.ReadWrite.All"],
    description: "Get the list of applications in this organization",
    version: "v1.0",
  },
  {
    id: "teams-list",
    name: "List Teams",
    url: "https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')",
    method: "GET",
    category: "Teams",
    scopes: ["Team.ReadBasic.All", "Group.Read.All"],
    description: "List all teams in Microsoft Teams",
    version: "beta",
  },
];

const categories = [
  "All",
  "Users",
  "Groups",
  "Mail",
  "Calendar",
  "Files",
  "Applications",
  "Teams",
];
const scopes = [
  "User.Read",
  "User.Read.All",
  "User.ReadWrite.All",
  "Group.Read.All",
  "Group.ReadWrite.All",
  "Mail.Read",
  "Mail.ReadWrite",
  "Calendars.Read",
  "Calendars.ReadWrite",
  "Files.Read",
  "Files.ReadWrite",
  "Files.Read.All",
  "Application.Read.All",
  "Application.ReadWrite.All",
  "Team.ReadBasic.All",
  "Directory.Read.All",
];

export default function Index() {
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("All");
  const [selectedScopes, setSelectedScopes] = useState<string[]>([]);
  const [selectedVersion, setSelectedVersion] = useState("All");
  const [selectedMethod, setSelectedMethod] = useState("All");
  const [copiedEndpoint, setCopiedEndpoint] = useState<string | null>(null);

  const filteredEndpoints = endpoints.filter((endpoint) => {
    const matchesSearch =
      endpoint.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      endpoint.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
      endpoint.url.toLowerCase().includes(searchQuery.toLowerCase());

    const matchesCategory =
      selectedCategory === "All" || endpoint.category === selectedCategory;
    const matchesVersion =
      selectedVersion === "All" || endpoint.version === selectedVersion;
    const matchesMethod =
      selectedMethod === "All" || endpoint.method === selectedMethod;

    const matchesScopes =
      selectedScopes.length === 0 ||
      selectedScopes.some((scope) => endpoint.scopes.includes(scope));

    return (
      matchesSearch &&
      matchesCategory &&
      matchesVersion &&
      matchesMethod &&
      matchesScopes
    );
  });

  const copyToClipboard = async (text: string, endpointId: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedEndpoint(endpointId);
      setTimeout(() => setCopiedEndpoint(null), 2000);
    } catch (err) {
      console.error("Failed to copy: ", err);
    }
  };

  const getMethodColor = (method: string) => {
    switch (method) {
      case "GET":
        return "bg-microsoft-blue text-white";
      case "POST":
        return "bg-microsoft-green text-white";
      case "PUT":
        return "bg-microsoft-orange text-white";
      case "DELETE":
        return "bg-microsoft-red text-white";
      default:
        return "bg-gray-500 text-white";
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50 dark:from-slate-900 dark:to-slate-800">
      {/* Header */}
      <div className="bg-white/80 dark:bg-slate-900/80 backdrop-blur-sm border-b sticky top-0 z-10">
        <div className="container mx-auto px-4 py-6">
          <div className="flex items-center space-x-4 mb-6">
            <div className="w-8 h-8 bg-graph-primary rounded-lg flex items-center justify-center">
              <div className="w-4 h-4 bg-white rounded-sm"></div>
            </div>
            <div>
              <h1 className="text-3xl font-bold text-slate-900 dark:text-white">
                Microsoft Graph API Explorer
              </h1>
              <p className="text-slate-600 dark:text-slate-300">
                Explore and filter Microsoft Graph endpoints for MSAL
                integration
              </p>
            </div>
          </div>

          {/* Search and Filters */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <div className="lg:col-span-4">
              <div className="relative">
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-slate-400 w-4 h-4" />
                <Input
                  placeholder="Search endpoints..."
                  className="pl-10"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
              </div>
            </div>

            <div className="lg:col-span-2">
              <Select
                value={selectedCategory}
                onValueChange={setSelectedCategory}
              >
                <SelectTrigger>
                  <SelectValue placeholder="Category" />
                </SelectTrigger>
                <SelectContent>
                  {categories.map((category) => (
                    <SelectItem key={category} value={category}>
                      {category}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="lg:col-span-2">
              <Select
                value={selectedVersion}
                onValueChange={setSelectedVersion}
              >
                <SelectTrigger>
                  <SelectValue placeholder="API Version" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="All">All Versions</SelectItem>
                  <SelectItem value="v1.0">v1.0</SelectItem>
                  <SelectItem value="beta">Beta</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div className="lg:col-span-2">
              <Select value={selectedMethod} onValueChange={setSelectedMethod}>
                <SelectTrigger>
                  <SelectValue placeholder="HTTP Method" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="All">All Methods</SelectItem>
                  <SelectItem value="GET">GET</SelectItem>
                  <SelectItem value="POST">POST</SelectItem>
                  <SelectItem value="PUT">PUT</SelectItem>
                  <SelectItem value="DELETE">DELETE</SelectItem>
                </SelectContent>
              </Select>
            </div>

            <div className="lg:col-span-2">
              <Button variant="outline" className="w-full">
                <Filter className="w-4 h-4 mr-2" />
                Advanced Filters
              </Button>
            </div>
          </div>
        </div>
      </div>

      <div className="container mx-auto px-4 py-8">
        {/* Clive Document Downloader Section */}
        <div className="mb-12">
          <CliveDownloader />
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          {/* Scope Filter Sidebar */}
          <div className="lg:col-span-1">
            <Card className="sticky top-24">
              <CardHeader>
                <CardTitle className="text-lg">Required Scopes</CardTitle>
                <CardDescription>
                  Filter by MSAL permission scopes
                </CardDescription>
              </CardHeader>
              <CardContent className="space-y-3">
                {scopes.map((scope) => (
                  <div key={scope} className="flex items-center space-x-2">
                    <Checkbox
                      id={scope}
                      checked={selectedScopes.includes(scope)}
                      onCheckedChange={(checked) => {
                        if (checked) {
                          setSelectedScopes([...selectedScopes, scope]);
                        } else {
                          setSelectedScopes(
                            selectedScopes.filter((s) => s !== scope),
                          );
                        }
                      }}
                    />
                    <label
                      htmlFor={scope}
                      className="text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 cursor-pointer"
                    >
                      {scope}
                    </label>
                  </div>
                ))}
                {selectedScopes.length > 0 && (
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => setSelectedScopes([])}
                    className="w-full mt-4"
                  >
                    Clear All
                  </Button>
                )}
              </CardContent>
            </Card>
          </div>

          {/* Endpoints List */}
          <div className="lg:col-span-3">
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-xl font-semibold text-slate-900 dark:text-white">
                API Endpoints ({filteredEndpoints.length})
              </h2>
            </div>

            <div className="space-y-4">
              {filteredEndpoints.map((endpoint) => (
                <Card
                  key={endpoint.id}
                  className="hover:shadow-md transition-shadow"
                >
                  <CardContent className="p-6">
                    <div className="flex items-start justify-between mb-4">
                      <div className="flex items-center space-x-3">
                        <Badge
                          className={cn(
                            "px-2 py-1 text-xs font-medium",
                            getMethodColor(endpoint.method),
                          )}
                        >
                          {endpoint.method}
                        </Badge>
                        <Badge
                          variant={
                            endpoint.version === "beta"
                              ? "destructive"
                              : "secondary"
                          }
                        >
                          {endpoint.version}
                        </Badge>
                        <Badge variant="outline">{endpoint.category}</Badge>
                      </div>
                      <div className="flex space-x-2">
                        <Button
                          variant="ghost"
                          size="sm"
                          onClick={() =>
                            copyToClipboard(endpoint.url, endpoint.id)
                          }
                        >
                          {copiedEndpoint === endpoint.id ? (
                            <Check className="w-4 h-4 text-green-500" />
                          ) : (
                            <Copy className="w-4 h-4" />
                          )}
                        </Button>
                        <Button variant="ghost" size="sm" asChild>
                          <a
                            href={`https://docs.microsoft.com/en-us/graph/api/`}
                            target="_blank"
                            rel="noopener noreferrer"
                          >
                            <ExternalLink className="w-4 h-4" />
                          </a>
                        </Button>
                      </div>
                    </div>

                    <h3 className="text-lg font-semibold text-slate-900 dark:text-white mb-2">
                      {endpoint.name}
                    </h3>

                    <p className="text-slate-600 dark:text-slate-300 mb-4">
                      {endpoint.description}
                    </p>

                    <div className="bg-slate-50 dark:bg-slate-800 rounded-lg p-3 mb-4">
                      <code className="text-sm text-slate-700 dark:text-slate-300 break-all">
                        {endpoint.url}
                      </code>
                    </div>

                    <div>
                      <h4 className="text-sm font-medium text-slate-900 dark:text-white mb-2">
                        Required Scopes:
                      </h4>
                      <div className="flex flex-wrap gap-2">
                        {endpoint.scopes.map((scope) => (
                          <Badge
                            key={scope}
                            variant="secondary"
                            className="text-xs"
                          >
                            {scope}
                          </Badge>
                        ))}
                      </div>
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>

            {filteredEndpoints.length === 0 && (
              <Card>
                <CardContent className="p-8 text-center">
                  <div className="text-slate-400 dark:text-slate-500 mb-2">
                    <Filter className="w-12 h-12 mx-auto" />
                  </div>
                  <h3 className="text-lg font-medium text-slate-900 dark:text-white mb-2">
                    No endpoints found
                  </h3>
                  <p className="text-slate-600 dark:text-slate-300">
                    Try adjusting your search criteria or filters.
                  </p>
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
