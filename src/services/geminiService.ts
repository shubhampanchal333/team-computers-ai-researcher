import { GoogleGenAI, Type, ThinkingLevel } from "@google/genai";

let ai: GoogleGenAI;
const getAI = () => {
  if (!ai) {
    const key = process.env.GEMINI_API_KEY as string;
    if (!key || key === 'undefined') {
      throw new Error("GEMINI_API_KEY is missing. Please add it to your Vercel environment variables.");
    }
    ai = new GoogleGenAI({ apiKey: key });
  }
  return ai;
};

export interface CompanyResearch {
  name: string;
  domain: string;
  logoUrl: string;
  isDeadpooled: boolean;
  arr: string;
  founders: Array<{
    name: string;
    role: string;
    linkedin: string;
    email: string;
    phone: string;
  }>;
  leadership: Array<{
    name: string;
    title: string;
    linkedin: string;
    focus: string;
    reportsTo?: string;
  }>;
  orgChart: {
    ceo: string;
    structure: Array<{
      from: string;
      to: string;
      relationship: string;
    }>;
  };
  funding: {
    isFunded: boolean;
    round: string;
    totalAmount?: string;
    lastFundingDate?: string;
    investors?: string[];
  };
  techStack: {
    cloud: string[];
    workspace: string[];
    security: string[];
    businessApps: string[];
    other: string[];
  };
  techStackSpend: {
    cloud: { subtotal: string; items: Array<{ name: string; spend: string }> };
    workspace: { subtotal: string; items: Array<{ name: string; spend: string }> };
    security: { subtotal: string; items: Array<{ name: string; spend: string }> };
    businessApps: { subtotal: string; items: Array<{ name: string; spend: string }> };
    other: { subtotal: string; items: Array<{ name: string; spend: string }> };
  };
  teamSize: string;
  customerBase: string;
  itInfrastructure: {
    endpoints: {
      laptops: string;
      desktops: string;
      mobile: string;
    };
    compute: {
      servers: string;
      storage: string;
      networking: string;
    };
    currentOEMs: string[];
    interestedOEMs: string[];
    fullStackAudit: string;
    oemAnalysis: string;
  };
  painPoints: Array<{
    point: string;
    impact: string;
  }>;
  serviceInterest: {
    infrastructure: number;
    cloud: number;
    managedServices: number;
    digitalTransformation: number;
    businessApps: number;
  };
  teamComputersPitch: {
    useCases: Array<{ 
      title: string; 
      description: string; 
      keyTakeaway: string; 
      serviceCategory: string;
      roiProjection: string;
      implementationTimeline: string;
    }>;
    strategicAdvantage: string;
    callScript: {
      introduction: string;
      discoveryQuestions: string[];
      valueProposition: string;
      handlingObjections: string[];
      closing: string;
    };
    upsellCrossSellOpportunities: Array<{
      title: string;
      description: string;
      targetService: string;
    }>;
  };
  competitiveLandscape: Array<{
    competitorName: string;
    serviceBenefit: string;
    teamComputersEdge: string;
    techStack: string[];
    services: string[];
  }>;
  teamComputersImpact: Array<{
    benefit: string;
    description: string;
    revenueImpact: string;
  }>;
  pitchDeck: {
    introduction: string;
    problemStatement: string;
    solution: string;
    marketAnalysis: string;
    competitiveAdvantage: string;
    proposedSolution: string;
    caseStudies: string[];
    callToAction: string;
  };
  sources: Array<{
    title: string;
    url: string;
    type: string;
  }>;
}

export interface ChatMessage {
  role: 'user' | 'model';
  content: string;
}

export interface CompanySummary {
  name: string;
  domain: string;
  industry: string;
  location: string;
  teamSize: string;
  description: string;
  estimatedMonthlyTechSpend: string;
}

export async function discoverCompanies(prompt: string): Promise<CompanySummary[]> {
  const response = await getAI().models.generateContent({
    model: "gemini-3-flash-preview",
    contents: `Quickly discover companies based on: "${prompt}". 
    Provide a list of up to 20 high-potential companies.
    For each, provide name, domain, industry, location, team size, description, and an estimatedMonthlyTechSpend (e.g. "$50,000").
    Focus on IT infrastructure relevance for Team Computers.`,
    config: {
      tools: [{ googleSearch: {} }],
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            domain: { type: Type.STRING },
            industry: { type: Type.STRING },
            location: { type: Type.STRING },
            teamSize: { type: Type.STRING },
            description: { type: Type.STRING },
            estimatedMonthlyTechSpend: { type: Type.STRING },
          },
          required: ["name", "domain", "industry", "location", "teamSize", "description", "estimatedMonthlyTechSpend"],
        },
      },
    },
  });

  try {
    return JSON.parse(response.text);
  } catch (e) {
    console.error("Failed to parse discovery response:", response.text);
    return [];
  }
}

export async function bulkSummarize(items: string[]): Promise<CompanySummary[]> {
  const response = await getAI().models.generateContent({
    model: "gemini-3-flash-preview",
    contents: `Carefully summarize these ${items.length} companies: ${items.join(", ")}.
    Provide a high-precision summary for EACH company. Return a total of ${items.length} objects.
    For each, provide name, domain, industry, location, team size, a brief business description, and an estimatedMonthlyTechSpend (e.g. "$15,000").`,
    config: {
      tools: [{ googleSearch: {} }],
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            name: { type: Type.STRING },
            domain: { type: Type.STRING },
            industry: { type: Type.STRING },
            location: { type: Type.STRING },
            teamSize: { type: Type.STRING },
            description: { type: Type.STRING },
            estimatedMonthlyTechSpend: { type: Type.STRING },
          },
          required: ["name", "domain", "industry", "location", "teamSize", "description", "estimatedMonthlyTechSpend"],
        },
      },
    },
  });

  try {
    return JSON.parse(response.text);
  } catch (e) {
    console.error("Failed to parse bulk summary response:", response.text);
    return [];
  }
}

export async function researchCompany(query: string): Promise<CompanyResearch> {
  const response = await getAI().models.generateContent({
    model: "gemini-3-flash-preview",
    contents: `EXTREME SPEED SCAN: ${query}. 
    Immediate structured intelligence for Team Computers.
    
    REQUIRED (Minimal & Fast):
    1. Basics: Status, logo, name, domain.
    2. Financial: ARR Range.
    3. Hierarchy: CEO, 2-3 Key CXOs & Reporting (Org Chart).
    4. Tech: Primary Cloud/Workspace/Security stack + category spend.
    5. Infra: Endpoints/Compute/OEM counts.
    6. Strategy: 2 Pain Points, 3 ROI Use-Cases, 1 Upsell.
    7. Sales: Brief Call Script, 8-slide Pitch Deck Outline.
    8. Competitive: 2 direct rivals + TC edge.
    9. Refs: 3-5 Source URLs.`,
    config: {
      tools: [{ googleSearch: {} }],
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          name: { type: Type.STRING },
          domain: { type: Type.STRING },
          logoUrl: { type: Type.STRING },
          isDeadpooled: { type: Type.BOOLEAN },
          arr: { type: Type.STRING },
          founders: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                name: { type: Type.STRING },
                role: { type: Type.STRING },
                linkedin: { type: Type.STRING },
                email: { type: Type.STRING },
                phone: { type: Type.STRING },
              },
              required: ["name", "role", "linkedin", "email", "phone"],
            },
          },
          leadership: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                name: { type: Type.STRING },
                title: { type: Type.STRING },
                linkedin: { type: Type.STRING },
                focus: { type: Type.STRING },
                reportsTo: { type: Type.STRING },
              },
              required: ["name", "title", "linkedin", "focus"],
            },
          },
          orgChart: {
            type: Type.OBJECT,
            properties: {
              ceo: { type: Type.STRING },
              structure: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    from: { type: Type.STRING },
                    to: { type: Type.STRING },
                    relationship: { type: Type.STRING },
                  },
                  required: ["from", "to", "relationship"],
                },
              },
            },
            required: ["ceo", "structure"],
          },
          funding: {
            type: Type.OBJECT,
            properties: {
              isFunded: { type: Type.BOOLEAN },
              round: { type: Type.STRING },
              totalAmount: { type: Type.STRING },
              lastFundingDate: { type: Type.STRING },
              investors: { type: Type.ARRAY, items: { type: Type.STRING } },
            },
            required: ["isFunded", "round"],
          },
          techStack: {
            type: Type.OBJECT,
            properties: {
              cloud: { type: Type.ARRAY, items: { type: Type.STRING } },
              workspace: { type: Type.ARRAY, items: { type: Type.STRING } },
              security: { type: Type.ARRAY, items: { type: Type.STRING } },
              businessApps: { type: Type.ARRAY, items: { type: Type.STRING } },
              other: { type: Type.ARRAY, items: { type: Type.STRING } },
            },
            required: ["cloud", "workspace", "security", "businessApps", "other"],
          },
          techStackSpend: {
            type: Type.OBJECT,
            properties: {
              cloud: {
                type: Type.OBJECT,
                properties: {
                  subtotal: { type: Type.STRING },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        spend: { type: Type.STRING },
                      },
                      required: ["name", "spend"],
                    },
                  },
                },
                required: ["subtotal", "items"],
              },
              workspace: {
                type: Type.OBJECT,
                properties: {
                  subtotal: { type: Type.STRING },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        spend: { type: Type.STRING },
                      },
                      required: ["name", "spend"],
                    },
                  },
                },
                required: ["subtotal", "items"],
              },
              security: {
                type: Type.OBJECT,
                properties: {
                  subtotal: { type: Type.STRING },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        spend: { type: Type.STRING },
                      },
                      required: ["name", "spend"],
                    },
                  },
                },
                required: ["subtotal", "items"],
              },
              businessApps: {
                type: Type.OBJECT,
                properties: {
                  subtotal: { type: Type.STRING },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        spend: { type: Type.STRING },
                      },
                      required: ["name", "spend"],
                    },
                  },
                },
                required: ["subtotal", "items"],
              },
              other: {
                type: Type.OBJECT,
                properties: {
                  subtotal: { type: Type.STRING },
                  items: {
                    type: Type.ARRAY,
                    items: {
                      type: Type.OBJECT,
                      properties: {
                        name: { type: Type.STRING },
                        spend: { type: Type.STRING },
                      },
                      required: ["name", "spend"],
                    },
                  },
                },
                required: ["subtotal", "items"],
              },
            },
            required: ["cloud", "workspace", "security", "businessApps", "other"],
          },
          teamSize: { type: Type.STRING },
          customerBase: { type: Type.STRING },
          itInfrastructure: {
            type: Type.OBJECT,
            properties: {
              endpoints: {
                type: Type.OBJECT,
                properties: {
                  laptops: { type: Type.STRING },
                  desktops: { type: Type.STRING },
                  mobile: { type: Type.STRING },
                },
                required: ["laptops", "desktops", "mobile"],
              },
              compute: {
                type: Type.OBJECT,
                properties: {
                  servers: { type: Type.STRING },
                  storage: { type: Type.STRING },
                  networking: { type: Type.STRING },
                },
                required: ["servers", "storage", "networking"],
              },
              currentOEMs: { type: Type.ARRAY, items: { type: Type.STRING } },
              interestedOEMs: { type: Type.ARRAY, items: { type: Type.STRING } },
              fullStackAudit: { type: Type.STRING },
              oemAnalysis: { type: Type.STRING },
            },
            required: ["endpoints", "compute", "currentOEMs", "interestedOEMs", "fullStackAudit", "oemAnalysis"],
          },
          painPoints: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                point: { type: Type.STRING },
                impact: { type: Type.STRING },
              },
              required: ["point", "impact"],
            },
          },
          serviceInterest: {
            type: Type.OBJECT,
            properties: {
              infrastructure: { type: Type.NUMBER },
              cloud: { type: Type.NUMBER },
              managedServices: { type: Type.NUMBER },
              digitalTransformation: { type: Type.NUMBER },
              businessApps: { type: Type.NUMBER },
            },
            required: ["infrastructure", "cloud", "managedServices", "digitalTransformation", "businessApps"],
          },
          teamComputersPitch: {
            type: Type.OBJECT,
            properties: {
              useCases: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: { type: Type.STRING },
                    description: { type: Type.STRING },
                    keyTakeaway: { type: Type.STRING, description: "A brief, impactful summary of the benefit." },
                    serviceCategory: { type: Type.STRING },
                    roiProjection: { type: Type.STRING },
                    implementationTimeline: { type: Type.STRING },
                  },
                  required: ["title", "description", "keyTakeaway", "serviceCategory", "roiProjection", "implementationTimeline"],
                },
              },
              strategicAdvantage: { type: Type.STRING },
              callScript: {
                type: Type.OBJECT,
                properties: {
                  introduction: { type: Type.STRING },
                  discoveryQuestions: { type: Type.ARRAY, items: { type: Type.STRING } },
                  valueProposition: { type: Type.STRING },
                  handlingObjections: { type: Type.ARRAY, items: { type: Type.STRING } },
                  closing: { type: Type.STRING },
                },
                required: ["introduction", "discoveryQuestions", "valueProposition", "handlingObjections", "closing"],
              },
              upsellCrossSellOpportunities: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: { type: Type.STRING },
                    description: { type: Type.STRING },
                    targetService: { type: Type.STRING },
                  },
                  required: ["title", "description", "targetService"],
                },
              },
            },
            required: ["useCases", "strategicAdvantage", "callScript", "upsellCrossSellOpportunities"],
          },
          competitiveLandscape: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                competitorName: { type: Type.STRING },
                serviceBenefit: { type: Type.STRING },
                teamComputersEdge: { type: Type.STRING },
                techStack: { type: Type.ARRAY, items: { type: Type.STRING } },
                services: { type: Type.ARRAY, items: { type: Type.STRING } },
              },
              required: ["competitorName", "serviceBenefit", "teamComputersEdge", "techStack", "services"],
            },
          },
          teamComputersImpact: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                benefit: { type: Type.STRING },
                description: { type: Type.STRING },
                revenueImpact: { type: Type.STRING },
              },
              required: ["benefit", "description", "revenueImpact"],
            },
          },
          pitchDeck: {
            type: Type.OBJECT,
            properties: {
              introduction: { type: Type.STRING },
              problemStatement: { type: Type.STRING },
              solution: { type: Type.STRING },
              marketAnalysis: { type: Type.STRING },
              competitiveAdvantage: { type: Type.STRING },
              proposedSolution: { type: Type.STRING },
              caseStudies: { type: Type.ARRAY, items: { type: Type.STRING } },
              callToAction: { type: Type.STRING },
            },
            required: ["introduction", "problemStatement", "solution", "marketAnalysis", "competitiveAdvantage", "proposedSolution", "caseStudies", "callToAction"],
          },
          sources: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                title: { type: Type.STRING },
                url: { type: Type.STRING },
                type: { type: Type.STRING },
              },
              required: ["title", "url", "type"],
            },
          },
        },
        required: ["name", "domain", "logoUrl", "isDeadpooled", "arr", "founders", "leadership", "orgChart", "funding", "techStack", "techStackSpend", "teamSize", "customerBase", "itInfrastructure", "painPoints", "serviceInterest", "teamComputersPitch", "competitiveLandscape", "teamComputersImpact", "pitchDeck", "sources"],
      },
    },
  });

  try {
    return JSON.parse(response.text);
  } catch (e) {
    console.error("Failed to parse response:", response.text);
    throw new Error("Failed to research company. Please try again.");
  }
}

export async function getChatResponse(history: ChatMessage[], context: CompanyResearch): Promise<string> {
  const response = await getAI().models.generateContent({
    model: "gemini-3-flash-preview",
    contents: `You are the "Team Computers Intelligence Bot". You have access to a specific account report for the company "${context.name}".
    Your goal is to answer follow-up questions from sales representatives about this company, its competitors, and the strategic pitch.
    Use ONLY the provided context and your internal knowledge to provide helpful, strategic advice.
    If referencing specific data from the report (like stack spend or leadership), point it out clearly.
    Keep responses concise, professional, and focused on helping the sales team close the deal.
    
    CONTEXT DATA:
    ${JSON.stringify(context, null, 2)}
    
    CONVERSATION HISTORY:
    ${history.map(m => `${m.role.toUpperCase()}: ${m.content}`).join('\n')}
    
    USER: ${history[history.length - 1].content}`,
    config: {
      tools: [{ googleSearch: {} }],
      thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
    }
  });

  return response.text;
}
