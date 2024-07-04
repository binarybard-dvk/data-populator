"use client";

import { zodResolver } from "@hookform/resolvers/zod";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Form,
  FormControl,
  FormField,
  FormItem,
  FormLabel,
  FormMessage,
} from "@/components/ui/form";
import { useForm } from "react-hook-form";
import Link from "next/link";
import { z } from "zod";
import { useRouter } from "next/navigation";
import { emailLogin, oAuthLogin } from "@/app/actions";
import { toast } from "@hookform/resolvers/zod";

export const authSchema = z.object({
  email: z.string().min(2, "Email is required"),
  // password: z.string().min(8),
});

export default function SignIn() {
  const router = useRouter();
  const form = useForm<z.infer<typeof authSchema>>({
    resolver: zodResolver(authSchema),
    defaultValues: {
      email: "",
      // password: '',
    },
  });

  return (
    <div className="m-auto max-w-sm space-y-6">
      <div className="space-y-2 text-center">
        <h1 className="text-3xl font-bold">Welcome back!</h1>
        <p className="text-muted-foreground">
          Enter your email below to login to your account.
        </p>
      </div>

      <Form {...form}>
        <form
          onSubmit={form.handleSubmit((values) => emailLogin(values.email))}
          className="space-y-4"
        >
          <FormField
            control={form.control}
            name="email"
            rules={{
              required: "Email is required",
              pattern: {
                value: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
                message: "Please enter a valid email address",
              },
            }}
            render={({ field }) => (
              <FormItem>
                <FormLabel>Email</FormLabel>
                <FormControl>
                  <Input placeholder="data@crafter.com" {...field} />
                </FormControl>
                <FormMessage />
              </FormItem>
            )}
          />
          {/* <FormField
				control={form.control}
				name='password'
				rules={{
					required: 'Password is required',
					minLength: {
						value: 8,
						message: 'Password must be at least 8 characters',
					},
					pattern: {
						value:
							/^(?=.*[a-z])(?=.*[A-Z])(?=.*[0-9])(?=.*[!@#$%^&*])(?=.{8,})/,
						message:
							'Password must contain at least 8 characters, one letter, one number and one special character',
					},
				}}
				render={({ field }) => (
					<FormItem>
						<FormLabel>Password</FormLabel>
						<FormControl>
							<Input
								type='password'
								placeholder='super secret ✨'
								{...field}
							/>
						</FormControl>
						<FormMessage />
					</FormItem>
				)}
			/> */}
          <Button
            type="submit"
            className="w-full"
            disabled={form.formState.isSubmitting}
          >
            Continue with Email
          </Button>
        </form>
      </Form>

      <div className="my-6 flex items-center justify-center">
        <Separator className="w-[40%]"></Separator>
        <span className="mx-4">or</span>
        <Separator className="w-[40%]"></Separator>
      </div>
      <div className="space-y-2"></div>
      <div className="space-y-4">
        <Button
          variant="outline"
          className="w-full"
          onClick={async () => await oAuthLogin("google")}
        >
          Continue with Google
        </Button>
        <Button
          variant="outline"
          className="w-full"
          onClick={async () => await oAuthLogin("github")}
        >
          Continue with GitHub
        </Button>
      </div>

      <p className="px-8 text-center text-sm text-muted-foreground">
        By clicking continue, you agree to our{" "}
        <Link
          href="/terms"
          className="underline underline-offset-4 hover:text-primary"
        >
          Terms of Service
        </Link>{" "}
        and{" "}
        <Link
          href="/privacy"
          className="underline underline-offset-4 hover:text-primary"
        >
          Privacy Policy
        </Link>
        .
      </p>
    </div>
  );
}
