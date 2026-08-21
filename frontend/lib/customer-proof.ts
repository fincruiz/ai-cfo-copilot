export type ApprovedCustomerProof = {
  customer_name: string;
  quote: string;
  role?: string;
  outcome?: string;
  permission_reference: string;
};

// Stage 9.9 proof-safety rule:
// Add entries only after the customer has explicitly approved public use.
// Keeping this empty prevents fabricated testimonials or case studies from shipping.
export const approvedCustomerProof: ApprovedCustomerProof[] = [];
