import { createFileRoute } from "@tanstack/react-router";
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
import { useEffect, useState } from "react";

export const Route = createFileRoute("/auth-callback")({
  component: AuthCallback,
});

function AuthCallback() {
  const [message, setMessage] = useState("Concluindo login…");

  useEffect(() => {
    broadcastResponseToMainFrame().catch((error: Error) => {
      console.error("Erro ao concluir autenticação no popup:", error);
      setMessage("Não foi possível concluir o login. Feche esta janela e tente novamente.");
    });
  }, []);

  return (
    <main className="grid min-h-screen place-items-center bg-background px-6 text-center text-sm text-muted-foreground">
      {message}
    </main>
  );
}
