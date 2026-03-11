import { createRootRoute, createRoute, createRouter, Outlet } from "@tanstack/react-router";
import App from "./App";

function RootLayout() {
  return <Outlet />;
}

const rootRoute = createRootRoute({
  component: RootLayout,
});

const indexRoute = createRoute({
  getParentRoute: () => rootRoute,
  path: "/",
  component: () => <App />,
});

const threadRoute = createRoute({
  getParentRoute: () => rootRoute,
  path: "/$threadId",
  component: () => <App />,
});

const routeTree = rootRoute.addChildren([indexRoute, threadRoute]);

export const router = createRouter({
  routeTree,
  defaultPreload: "intent",
});

declare module "@tanstack/react-router" {
  interface Register {
    router: typeof router;
  }
}
