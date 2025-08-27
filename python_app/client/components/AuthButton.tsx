import { useState } from "react";
import { useMsal, useAccount, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "@/config/msalConfig";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  User,
  LogIn,
  LogOut,
  CheckCircle,
  AlertCircle,
  Loader2,
} from "lucide-react";

interface AuthButtonProps {
  onAuthSuccess?: () => void;
}

export default function AuthButton({ onAuthSuccess }: AuthButtonProps) {
  const { instance, accounts } = useMsal();
  const account = useAccount(accounts[0] || {});
  const isAuthenticated = useIsAuthenticated();
  const [isLoading, setIsLoading] = useState(false);

  const handleLogin = async () => {
    setIsLoading(true);
    try {
      const response = await instance.loginPopup(loginRequest);
      console.log("Login successful:", response);

      // Call the callback when authentication is successful
      if (onAuthSuccess) {
        setTimeout(() => {
          onAuthSuccess();
        }, 1000); // Small delay to let the UI update
      }
    } catch (error) {
      console.error("Login failed:", error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = () => {
    instance.logoutPopup({
      postLogoutRedirectUri: window.location.origin,
    });
  };

  if (isAuthenticated && account) {
    return (
      <Card className="w-full max-w-md">
        <CardHeader className="pb-3">
          <CardTitle className="flex items-center gap-2 text-sm">
            <CheckCircle className="w-4 h-4 text-green-500" />
            Authenticated
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-microsoft-blue rounded-full flex items-center justify-center">
              <User className="w-4 h-4 text-white" />
            </div>
            <div>
              <div className="font-medium text-sm">{account.name}</div>
              <div className="text-xs text-muted-foreground">
                {account.username}
              </div>
            </div>
          </div>

          <div className="flex flex-wrap gap-1">
            {loginRequest.scopes?.map((scope) => (
              <Badge key={scope} variant="secondary" className="text-xs">
                {scope}
              </Badge>
            ))}
          </div>

          <Button
            onClick={handleLogout}
            variant="outline"
            size="sm"
            className="w-full"
          >
            <LogOut className="w-4 h-4 mr-2" />
            Sign Out
          </Button>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card className="w-full max-w-md">
      <CardHeader className="pb-3">
        <CardTitle className="flex items-center gap-2 text-sm">
          <AlertCircle className="w-4 h-4 text-orange-500" />
          Microsoft Authentication
        </CardTitle>
        <CardDescription className="text-xs">
          Sign in with Microsoft to enable Clive document downloads
        </CardDescription>
      </CardHeader>
      <CardContent>
        <Button
          onClick={handleLogin}
          disabled={isLoading}
          className="w-full bg-microsoft-blue hover:bg-microsoft-blue/90"
        >
          {isLoading ? (
            <>
              <Loader2 className="w-4 h-4 mr-2 animate-spin" />
              Signing in...
            </>
          ) : (
            <>
              <LogIn className="w-4 h-4 mr-2" />
              Sign in with Microsoft
            </>
          )}
        </Button>
      </CardContent>
    </Card>
  );
}
