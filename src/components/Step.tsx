import { cn } from "@/lib/utils";

interface StepProps {
  step: number;
  title: string;
  className?: string;
}

export function Step({ step, title, className }: StepProps) {
  return (
    <div className={cn("flex items-center gap-4", className)}>
      <div className="flex h-10 w-10 items-center justify-center rounded-full border-2 border-primary bg-primary/10 text-primary font-bold">
        {step}
      </div>
      <h2 className="text-xl font-semibold tracking-tight">{title}</h2>
    </div>
  );
}