import { NextResponse } from "next/server";
import {
  isLocalTestCredentials,
  LOCAL_TEST_EMAIL,
  LOCAL_TEST_NAME,
} from "@/lib/local-auth";

export async function POST(request: Request) {
  try {
    const body = (await request.json()) as {
      id?: unknown;
      password?: unknown;
    };
    const id = String(body.id ?? "");
    const password = String(body.password ?? "");

    if (!isLocalTestCredentials(id, password)) {
      return NextResponse.json(
        { error: "Invalid ID or password." },
        { status: 401 },
      );
    }

    return NextResponse.json({
      email: LOCAL_TEST_EMAIL,
      name: LOCAL_TEST_NAME,
    });
  } catch {
    return NextResponse.json(
      { error: "Invalid login request." },
      { status: 400 },
    );
  }
}
