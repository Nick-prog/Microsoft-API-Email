import { useState, useRef } from "react";
import { useIsAuthenticated } from "@azure/msal-react";
import {
  Download,
  FileText,
  AlertCircle,
  CheckCircle,
  Loader2,
  Lock,
} from "lucide-react";
import AuthButton from "@/components/AuthButton";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Alert, AlertDescription } from "@/components/ui/alert";

interface CliveInfo {
  status: number;
  statusText: string;
  contentType: string | null;
  contentDisposition: string | null;
  contentLength: string | null;
  finalUrl: string;
}

export default function CliveDownloader() {
  const [cliveUrl, setCliveUrl] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [info, setInfo] = useState<CliveInfo | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [showDownloader, setShowDownloader] = useState(false);
  const isAuthenticated = useIsAuthenticated();
  const downloaderRef = useRef<HTMLDivElement>(null);

  const handleAuthSuccess = () => {
    setShowDownloader(true);
    // Scroll to and focus on the downloader section
    setTimeout(() => {
      if (downloaderRef.current) {
        downloaderRef.current.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
        // Focus on the URL input field
        const urlInput = downloaderRef.current.querySelector("input");
        if (urlInput) {
          urlInput.focus();
        }
      }
    }, 500);
  };

  const getCliveInfo = async () => {
    if (!cliveUrl.trim()) {
      setError("Please enter a Clive URL");
      return;
    }

    if (!cliveUrl.includes("clive.cloud")) {
      setError("Please enter a valid Clive URL (must contain 'clive.cloud')");
      return;
    }

    setIsLoading(true);
    setError(null);
    setSuccess(null);
    setInfo(null);

    try {
      const response = await fetch(
        `/api/clive/info?url=${encodeURIComponent(cliveUrl)}`,
      );
      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "Failed to get Clive info");
      }

      setInfo(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to get Clive info");
    } finally {
      setIsLoading(false);
    }
  };

  const downloadCliveDocument = async () => {
    if (!cliveUrl.trim()) {
      setError("Please enter a Clive URL");
      return;
    }

    setIsLoading(true);
    setError(null);
    setSuccess(null);

    try {
      const response = await fetch("/api/clive/download", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ url: cliveUrl }),
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Download failed");
      }

      // Get filename from Content-Disposition header or use default
      const contentDisposition = response.headers.get("content-disposition");
      let filename = "clive-document";

      if (contentDisposition) {
        const filenameMatch = contentDisposition.match(
          /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/,
        );
        if (filenameMatch && filenameMatch[1]) {
          filename = filenameMatch[1].replace(/['"]/g, "");
        }
      }

      // Create blob and download
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

      setSuccess(`Document "${filename}" downloaded successfully!`);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Download failed");
    } finally {
      setIsLoading(false);
    }
  };

  const isValidCliveUrl = cliveUrl.includes("clive.cloud");

  return (
    <div className="w-full max-w-4xl mx-auto space-y-6">
      {/* Authentication Section */}
      {!isAuthenticated && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <Lock className="w-5 h-5" />
              Authentication Required
            </CardTitle>
            <CardDescription>
              Please authenticate with Microsoft to enable Clive document
              downloads
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="flex justify-center">
              <AuthButton onAuthSuccess={handleAuthSuccess} />
            </div>
          </CardContent>
        </Card>
      )}

      {/* Clive Downloader Section - Show after authentication */}
      {(isAuthenticated || showDownloader) && (
        <Card ref={downloaderRef} className="border-2 border-microsoft-blue/20">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <FileText className="w-5 h-5" />
              Clive Document Downloader
              {isAuthenticated && (
                <CheckCircle className="w-4 h-4 text-green-500 ml-auto" />
              )}
            </CardTitle>
            <CardDescription>
              Download documents from Clive URLs that redirect to webpages
              instead of files
            </CardDescription>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="space-y-2">
              <label htmlFor="clive-url" className="text-sm font-medium">
                Clive URL
              </label>
              <Input
                id="clive-url"
                placeholder="http://url821.clive.cloud/ls/click?upn=..."
                value={cliveUrl}
                onChange={(e) => setCliveUrl(e.target.value)}
                className="font-mono text-sm"
              />
            </div>

            <div className="flex gap-2">
              <Button
                onClick={getCliveInfo}
                disabled={!isValidCliveUrl || isLoading}
                variant="outline"
                size="sm"
              >
                {isLoading ? (
                  <Loader2 className="w-4 h-4 animate-spin mr-2" />
                ) : (
                  <AlertCircle className="w-4 h-4 mr-2" />
                )}
                Inspect URL
              </Button>
              <Button
                onClick={downloadCliveDocument}
                disabled={!isValidCliveUrl || isLoading}
                className="bg-microsoft-blue hover:bg-microsoft-blue/90"
              >
                {isLoading ? (
                  <Loader2 className="w-4 h-4 animate-spin mr-2" />
                ) : (
                  <Download className="w-4 h-4 mr-2" />
                )}
                Download Document
              </Button>
            </div>

            {error && (
              <Alert variant="destructive">
                <AlertCircle className="w-4 h-4" />
                <AlertDescription>{error}</AlertDescription>
              </Alert>
            )}

            {success && (
              <Alert className="border-green-200 bg-green-50 text-green-800">
                <CheckCircle className="w-4 h-4" />
                <AlertDescription>{success}</AlertDescription>
              </Alert>
            )}

            {info && (
              <Card className="bg-slate-50 dark:bg-slate-800">
                <CardHeader>
                  <CardTitle className="text-sm">
                    URL Analysis Results
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-2 text-sm">
                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <span className="font-medium">Status:</span> {info.status}{" "}
                      {info.statusText}
                    </div>
                    <div>
                      <span className="font-medium">Content Type:</span>{" "}
                      <code className="text-xs">
                        {info.contentType || "N/A"}
                      </code>
                    </div>
                    <div className="col-span-2">
                      <span className="font-medium">Content Disposition:</span>{" "}
                      <code className="text-xs">
                        {info.contentDisposition || "N/A"}
                      </code>
                    </div>
                    <div>
                      <span className="font-medium">Size:</span>{" "}
                      {info.contentLength || "Unknown"}
                    </div>
                    <div className="col-span-2">
                      <span className="font-medium">Final URL:</span>{" "}
                      <code className="text-xs break-all">{info.finalUrl}</code>
                    </div>
                  </div>
                  {info.contentType?.includes("text/html") && (
                    <Alert className="mt-2">
                      <AlertCircle className="w-4 h-4" />
                      <AlertDescription className="text-xs">
                        This URL returns HTML (webpage) instead of a file. The
                        downloader will attempt to extract the actual download
                        link from the page.
                      </AlertDescription>
                    </Alert>
                  )}
                </CardContent>
              </Card>
            )}

            <div className="text-xs text-slate-500 space-y-1">
              <p>
                <strong>How it works:</strong> This tool fetches the Clive URL
                and attempts to:
              </p>
              <ul className="list-disc list-inside space-y-1 ml-2">
                <li>Follow redirects to find the actual download link</li>
                <li>
                  Extract download URLs from HTML if redirected to a webpage
                </li>
                <li>Serve the file with proper download headers</li>
                <li>Handle various document formats (PDF, DOCX, XLSX, etc.)</li>
              </ul>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Instructions for authenticated users */}
      {isAuthenticated && (
        <Card className="bg-blue-50 border-blue-200">
          <CardContent className="pt-6">
            <div className="flex items-start gap-3">
              <CheckCircle className="w-5 h-5 text-blue-600 mt-0.5" />
              <div className="text-sm text-blue-800">
                <p className="font-medium mb-1">
                  You're authenticated and ready!
                </p>
                <p>
                  Paste your Clive URL above and click "Download Document" to
                  bypass webpage redirects and download the actual file.
                </p>
              </div>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
